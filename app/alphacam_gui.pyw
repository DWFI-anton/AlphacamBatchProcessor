"""Diagnostic Python GUI for Alphacam furniture-part processing.

The script intentionally keeps each Alphacam operation small and explicit:
open ARD, extract features, detect Side 2, create a ROT copy, apply AutoStyle,
and optionally export NC. It is a replacement for the orchestration part of
the VBA macros; the actual hole generation remains Alphacam AutoStyle.
"""

from __future__ import annotations

import logging
import os
import queue
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


DEFAULT_PROGID = "Broutaps.Application"
DEFAULT_AUTOSTYLE = r"C:\ALPHACAM1\LICOMDIR\Styles\DRILL\AUTODRILL.ara"
DEFAULT_TEMPLATE = r"C:\ALPHACAM1\LICOMDIR\Templates\DEFAULT-PAZ.adt"
NC_OUTPUT_MODE = 0


class ProcessingError(RuntimeError):
    pass


@dataclass
class StageResult:
    stage: str
    message: str
    details: dict[str, Any]


def normalize_wp_name(value: str) -> str:
    return (
        str(value)
        .strip()
        .upper()
        .replace("_", "")
        .replace("-", "")
        .replace(" ", "")
    )


def is_side_name(value: str, side: int) -> bool:
    # В Automation.bas используется InStr("SIDE2"), а не строгое равенство.
    return f"SIDE{side}" in normalize_wp_name(value)


class AlphacamSession:
    def __init__(self, progid: str, log: logging.Logger):
        self.progid = progid
        self.log = log
        self.app: Any = None

    def connect(self) -> Any:
        if os.name != "nt":
            raise ProcessingError("Alphacam COM доступен только в Windows.")
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ProcessingError("Установите pywin32 из requirements.txt.") from exc

        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.Dispatch(self.progid)
            self.log.info("Connected to %s", self.progid)
            return self.app
        except Exception:
            pythoncom.CoUninitialize()
            raise

    def close(self) -> None:
        if self.app is not None:
            self.app = None
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass

    def active_drawing(self) -> Any:
        if self.app is None or self.app.ActiveDrawing is None:
            raise ProcessingError("Нет активного ARD в Alphacam.")
        return self.app.ActiveDrawing


class AlphacamProcessor:
    def __init__(self, session: AlphacamSession):
        self.session = session
        self.app = session.app
        self.log = session.log

    def run_stage(self, name: str, action: Callable[[], Any]) -> Any:
        self.log.info("START %s", name)
        try:
            result = action()
            self.log.info("DONE %s: %s", name, result)
            return result
        except Exception as exc:
            self.log.error("FAILED %s: %s", name, exc)
            self.log.error(traceback.format_exc())
            raise ProcessingError(f"{name}: {exc}") from exc

    def open_drawing(self, path: Path) -> Any:
        path = path.resolve()
        if not path.is_file():
            raise ProcessingError(f"Файл не найден: {path}")
        self.app.OpenDrawing(str(path))
        drawing = self.session.active_drawing()
        self.log.info("Opened %s", drawing.FullName)
        return drawing

    def get_solid(self, drawing: Any) -> Any:
        solid_interface = drawing.SolidInterface
        bodies = solid_interface.BodiesInDrawing(drawing)
        if int(bodies.Count) != 1:
            raise ProcessingError(
                f"Ожидалось одно solid body, получено: {bodies.Count}"
            )
        return bodies.Item(1)

    def solid_dimensions(self, drawing: Any) -> tuple[float, float, float]:
        # Повторяем Processing.bas: размеры берутся из первого SolidPart,
        # а не как разница экстентов SolidBody.
        parts = drawing.SolidParts
        if int(parts.Count) < 1:
            raise ProcessingError("В чертеже отсутствует SolidPart.")
        part = parts.Item(1)
        width = abs(float(part.MaxX))
        height = abs(float(part.MinY))
        thickness = abs(float(part.MaxZ))
        if min(width, height, thickness) <= 0:
            raise ProcessingError(
                f"Некорректные габариты: {width} x {height} x {thickness}"
            )
        return width, height, thickness

    def solid_body_dimensions(self, drawing: Any) -> tuple[float, float, float]:
        """Габариты SolidBody для повернутой ROT-копии.

        Это соответствует GetSolidBodyDimensionsForAutomation из Automation.bas.
        После поворота SolidPart.MinY может быть нулём, поэтому для ROT нужны
        именно разницы Min/Max.
        """
        body = self.get_solid(drawing)
        width = abs(float(body.MaxX) - float(body.MinX))
        height = abs(float(body.MaxY) - float(body.MinY))
        thickness = abs(float(body.MaxZ) - float(body.MinZ))
        if min(width, height, thickness) <= 0:
            raise ProcessingError(
                f"Некорректные габариты SolidBody: {width} x {height} x {thickness}"
            )
        return width, height, thickness

    def set_drawing_parameters(
        self, drawing: Any, dimensions: tuple[float, float, float]
    ) -> None:
        width, height, thickness = dimensions
        algebra = drawing.ParametricInterface.ActiveAlgebra
        self.set_algebra_parameter(algebra, "Width", width)
        self.set_algebra_parameter(algebra, "Height", height)
        self.set_algebra_parameter(algebra, "Thickness", thickness)
        evaluate = drawing.ParametricInterface.ActiveIgm.Evaluate
        if callable(evaluate):
            evaluate = evaluate()
        self.log.info(
            "Параметры чертежа обновлены и геометрия пересчитана: "
            "Width=%.3f Height=%.3f Thickness=%.3f Result=%s",
            width, height, thickness, evaluate,
        )
        if not evaluate:
            raise ProcessingError("Parametric evaluation failed.")

    def set_algebra_parameter(self, algebra: Any, name: str, value: float) -> None:
        """Set VBA-style Algebra.Parameter(name) = value through pywin32.

        Alphacam exposes Parameter as a parameterized COM property. Depending
        on whether pywin32 generated an early-bound wrapper, it appears either
        as a setter method or only through IDispatch::Invoke.
        """
        errors: list[str] = []

        # Это основной рабочий вариант из alphacam_tools_gui.py.
        try:
            algebra.Parameter(name, float(value))
            return
        except Exception as exc:
            errors.append(f"Parameter(name, value): {exc}")

        try:
            parameter = algebra.Parameter(name)
            if hasattr(parameter, "Value"):
                parameter.Value = float(value)
                return
        except Exception as exc:
            errors.append(f"Parameter(name).Value: {exc}")

        for method_name in ("SetParameter", "SetParam", "Set"):
            try:
                method = getattr(algebra, method_name)
                method(name, value)
                return
            except Exception as exc:
                errors.append(f"{method_name}: {exc}")

        try:
            import pythoncom

            ole = algebra._oleobj_
            dispid = ole.GetIDsOfNames("Parameter")
            if isinstance(dispid, tuple):
                dispid = dispid[0]

            # Alphacam's VBA syntax is Parameter(name) = value. In the
            # pywin32 IDispatch call the indexed name precedes the value.
            ole.Invoke(
                dispid,
                0,
                pythoncom.DISPATCH_PROPERTYPUT,
                0,
                name,
                value,
            )
            return
        except Exception as exc:
            errors.append(f"Parameter property put: {exc}")

        try:
            parameter_method = getattr(algebra, "Parameter")
            parameter_method(name, float(value))
            return
        except Exception as exc:
            errors.append(f"Parameter method setter: {exc}")

        raise ProcessingError(
            f"Не удалось записать параметр Alphacam {name!r}. "
            + " | ".join(errors)
        )

    def workplanes(self, drawing: Any) -> list[Any]:
        collection = drawing.WorkPlanes
        return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]

    def extract_workplanes(
        self,
        drawing: Any,
        target_side: int | None = None,
        exclude_side: int | None = None,
    ) -> list[dict[str, Any]]:
        solid_interface = drawing.SolidInterface
        extraction = solid_interface.SolidExtraction
        extraction.FindContours = False
        results: list[dict[str, Any]] = []

        # VBA получает FeatureCurrentPlaneOrientation из подключённой
        # библиотеки Alphacam, а pywin32 в этой установке enum не публикует.
        # В этом случае запускаем тот же extFeat из VBA и только затем
        # сравниваем WorkPlane до/после.
        orientation = self.current_plane_orientation()
        if orientation is None:
            vba_results = self.extract_workplanes_via_vba(
                drawing, target_side, exclude_side
            )
            if vba_results is not None:
                return vba_results

        for workplane in self.workplanes(drawing):
            name = str(workplane.Name)
            if target_side is not None and not is_side_name(name, target_side):
                continue
            if exclude_side is not None and is_side_name(name, exclude_side):
                continue

            before = int(workplane.Geometries.Count)
            drawing.SetWorkPlane(workplane)
            extraction.SetWorkPlane(workplane)

            if orientation is not None:
                extraction.Orientation = orientation
            solid_interface.AutoExtract(extraction)
            after = int(workplane.Geometries.Count)

            item = {
                "name": name,
                "before": before,
                "after": after,
                "created": max(0, after - before),
            }
            results.append(item)
            self.log.info("WP %s: %s", name, item)

        drawing.SetWorkPlane(None)
        return results

    def extract_workplanes_via_vba(
        self,
        drawing: Any,
        target_side: int | None = None,
        exclude_side: int | None = None,
    ) -> list[dict[str, Any]] | None:
        """Выполнить ExtractFeature.extFeat через Alphacam VBA.

        Это точная копия макросного пути: SetWorkPlane, SetWorkPlane для
        SolidExtraction, FeatureCurrentPlaneOrientation и AutoExtract.
        """
        workplanes = self.workplanes(drawing)
        before = [int(workplane.Geometries.Count) for workplane in workplanes]
        project = os.environ.get("ALPHACAM_VBA_PROJECT", "DWFI_PR").strip()
        selected = []
        for index, workplane in enumerate(workplanes):
            name = str(workplane.Name)
            if target_side is not None and not is_side_name(name, target_side):
                continue
            if exclude_side is not None and is_side_name(name, exclude_side):
                continue
            selected.append((index, workplane, name))

        last_error: Exception | None = None
        if target_side is not None or exclude_side is not None:
            macro = f"{project}.Automation.ExtractFeaturesFromNamedWorkPlane"
            for index, workplane, name in selected:
                try:
                    try:
                        self.app.Run(macro, name)
                    except Exception:
                        self.app.Run(f"{macro}('{name}')")
                    self.log.info("Извлечение WorkPlane выполнено через VBA: %s", name)
                except Exception as exc:
                    last_error = exc
                    break
            if last_error is not None:
                self.log.warning("VBA-извлечение выбранных WorkPlane недоступно: %s", last_error)
                return None
        else:
            macros = (
                f"{project}.ExtractFeature.extFeat",
                f"{project}.Automation.ExtractFeaturesAndCheckSide2",
            )
            for macro in macros:
                try:
                    self.app.Run(macro)
                    self.log.info("Извлечение WorkPlane выполнено через VBA: %s", macro)
                    break
                except Exception as exc:
                    last_error = exc
            else:
                self.log.warning("VBA-извлечение недоступно: %s", last_error)
                return None

        results: list[dict[str, Any]] = []
        try:
            for index, workplane in enumerate(workplanes):
                name = str(workplane.Name)
                if target_side is not None and not is_side_name(name, target_side):
                    continue
                if exclude_side is not None and is_side_name(name, exclude_side):
                    continue
                after = int(workplane.Geometries.Count)
                results.append(
                    {
                        "name": name,
                        "before": before[index],
                        "after": after,
                        "created": max(0, after - before[index]),
                    }
                )
                self.log.info("WP %s: %s", name, results[-1])
        finally:
            drawing.SetWorkPlane(None)
        return results

    def current_plane_orientation(self) -> int | None:
        """Return the Alphacam enum when the registered type library exposes it.

        Some Alphacam installations do not publish this VBA enum to pywin32.
        In that case SetWorkPlane already establishes the active plane and the
        property assignment is skipped instead of aborting the whole run.
        """
        try:
            import win32com.client

            value = getattr(
                win32com.client.constants,
                "FeatureCurrentPlaneOrientation",
                None,
            )
            if value is not None:
                return int(value)
        except Exception as exc:
            self.log.debug("FeatureCurrentPlaneOrientation lookup: %s", exc)

        raw = os.environ.get("ALPHACAM_FEATURE_ORIENTATION", "").strip()
        if raw:
            try:
                value = int(raw)
                self.log.info("Using ALPHACAM_FEATURE_ORIENTATION=%s", value)
                return value
            except ValueError:
                self.log.warning("Invalid ALPHACAM_FEATURE_ORIENTATION=%r", raw)

        if not getattr(self, "_orientation_warning_shown", False):
            self.log.warning(
                "FeatureCurrentPlaneOrientation не опубликована в COM; "
                "как в макросе, используется активный WorkPlane через "
                "SetWorkPlane и SolidExtraction.SetWorkPlane."
            )
            self._orientation_warning_shown = True
        return None

    def apply_autostyle(self, path: Path) -> None:
        if not path.is_file():
            raise ProcessingError(f"AutoStyle не найден: {path}")
        import win32com.client

        # В разных сборках Alphacam зарегистрирован разный ProgID. В текущей
        # установке AcamAddInsInterface.AddInsInterface отсутствует, поэтому
        # пробуем штатные варианты и затем интерфейс, опубликованный самим App.
        addins_interface = None
        progids = (
            "Alphacam.AddIns.Interface.AddInsInterface",
            "AcamAddInsInterface.AddInsInterface",
            "Alphacam.AddIns.AddInsInterface",
            "AcamAddIns.AddInsInterface",
            "AcamAddInsInterface.AddInsInterface.1",
        )
        errors: list[str] = []
        for progid in progids:
            try:
                addins_interface = win32com.client.Dispatch(progid)
                self.log.info("AutoStyle AddIns: подключён %s", progid)
                break
            except Exception as exc:
                errors.append(f"{progid}: {exc}")

        addins = None
        if addins_interface is not None:
            for method_name in ("GetAddInsInterface", "getAddInsInterface"):
                try:
                    method = getattr(addins_interface, method_name)
                    addins = method(self.app)
                    if addins is not None:
                        break
                except Exception as exc:
                    errors.append(f"{method_name}: {exc}")

        if addins is None:
            try:
                addins = self.app.AddIns
                self.log.info("AutoStyle AddIns: используется App.AddIns")
            except Exception as exc:
                errors.append(f"App.AddIns: {exc}")

        if addins is None:
            raise ProcessingError(
                "Не удалось получить AddIns interface. " + " | ".join(errors)
            )

        auto_style = None
        for attr_name in ("GetAutoStylesAddIn", "getAutoStylesAddIn", "AutoStyles"):
            try:
                value = getattr(addins, attr_name)
                auto_style = value() if callable(value) else value
                if auto_style is not None:
                    break
            except Exception as exc:
                errors.append(f"{attr_name}: {exc}")
        if auto_style is None:
            raise ProcessingError(
                "Не удалось получить AutoStyles AddIn. " + " | ".join(errors)
            )
        auto_style.Apply(str(path))

    def save_as(self, drawing: Any, path: Path) -> None:
        drawing.SaveAs(str(path.resolve()))
        if not path.is_file():
            raise ProcessingError(f"Alphacam не создал файл: {path}")

    def import_x_t(self, source_path: Path, template_path: Path) -> Any:
        """Создать новый ARD из X_T и оставить его открытым для настройки."""
        if not source_path.is_file():
            raise ProcessingError(f"X_T не найден: {source_path}")
        if not template_path.is_file():
            raise ProcessingError(f"Шаблон ARD не найден: {template_path}")
        self.app.New()
        self.app.OpenDrawing(str(template_path.resolve()))
        self.app.OpenExternalFile(str(source_path.resolve()), "Parasolid", False)
        return self.session.active_drawing()

    def capture_groove_via_vba(self, drawing: Any) -> None:
        """Дать оператору выбрать грань паза и сохранить параметры в ARD."""
        project = os.environ.get("ALPHACAM_VBA_PROJECT", "DWFI_PR").strip()
        imp_macros = (
            f"{project}.Paz.impPaz",
        )
        last_error: Exception | None = None
        for macro in imp_macros:
            try:
                self.app.Run(macro)
                last_error = None
                break
            except Exception as exc:
                last_error = exc
        if last_error is not None:
            raise ProcessingError(
                "Не удалось запустить выбор паза через VBA. "
                "Проверьте, что модуль Paz загружен в проект Alphacam: "
                f"{last_error}"
            )

        save_macros = (
            f"{project}.Automation.SaveCurrentGrooveToARD",
            f"{project}.GrooveStorage.SaveCurrentGrooveToARD",
        )
        last_error = None
        for macro in save_macros:
            try:
                self.app.Run(macro)
                last_error = None
                break
            except Exception as exc:
                last_error = exc
        if last_error is not None:
            raise ProcessingError(
                "Параметры паза рассчитаны, но не сохранены в ARD: "
                f"{last_error}"
            )

        # SaveCurrentGrooveToARD не возвращает статус. Проверяем параметр,
        # чтобы не продолжать автоматическую обработку после отмены выбора.
        try:
            algebra = drawing.ParametricInterface.ActiveAlgebra
            enabled = int(algebra.Parameter("GrooveEnable"))
        except Exception as exc:
            raise ProcessingError(
                "Не удалось проверить сохранение параметров паза в ARD: "
                f"{exc}"
            ) from exc
        if enabled == 0:
            raise ProcessingError("Грань паза не выбрана или паз не сохранён.")
        self.log.info("Параметры паза выбраны и сохранены в ARD")

    def rotate_imported_solid(self, drawing: Any, angle: float, axis: str = "Z") -> None:
        """Повернуть solid и привязать нужный дальний нижний левый угол к 0,0,0.

        Логика повторяет Module1.bas:
        для RotateSolid90CW после -90° используются -MinX, -MaxY, -MinZ.
        Для обратного поворота применяется симметричная привязка. Переворот
        выполняется вокруг глобальной оси Y, после чего нижняя поверхность и
        левый край также возвращаются в начало координат.
        """
        body = self.get_solid(drawing)
        center_x = (float(body.MinX) + float(body.MaxX)) / 2.0
        center_y = (float(body.MinY) + float(body.MaxY)) / 2.0
        center_z = (float(body.MinZ) + float(body.MaxZ)) / 2.0
        if axis.upper() == "Y":
            body.RotateG(float(angle), center_x, center_y, center_z,
                         center_x, center_y + 1.0, center_z)
        else:
            body.RotateG(float(angle), center_x, center_y, center_z,
                         center_x, center_y, center_z + 1.0)

        if axis.upper() == "Y":
            # В RotateSolid180Y.bas перемещение было заготовлено, но
            # закомментировано. Здесь оно включено, чтобы выполнить требование
            # привязки детали к 0,0,0 после переворота.
            move_x = -float(body.MinX)
            move_y = -float(body.MinY)
            move_z = -float(body.MinZ)
        else:
            # Точка привязки должна быть одной и той же для обоих направлений:
            # левый дальний нижний угол -> (0, 0, 0). В RotateSolid90CW.bas
            # это -MinX, -MaxY, -MinZ. Нельзя использовать -MaxX/-MinY для
            # +90°, иначе в нуле оказывается правый угол детали.
            move_x = -float(body.MinX)
            move_y = -float(body.MaxY)
            move_z = -float(body.MinZ)

        body.MoveG(move_x, move_y, move_z)
        try:
            drawing.RedrawShadedViews()
        except Exception:
            pass

    def rotate_180_z_and_normalize(self, drawing: Any) -> None:
        body = self.get_solid(drawing)
        center_x = (float(body.MinX) + float(body.MaxX)) / 2.0
        center_y = (float(body.MinY) + float(body.MaxY)) / 2.0
        center_z = (float(body.MinZ) + float(body.MaxZ)) / 2.0
        body.RotateG(
            180.0,
            center_x,
            center_y,
            center_z,
            center_x,
            center_y,
            center_z + 1.0,
        )
        # Для ROT дальняя сторона должна лежать в Y=0, поэтому вся деталь
        # оказывается в отрицательной области Y: используется -MaxY.
        body.MoveG(-float(body.MinX), -float(body.MaxY), -float(body.MinZ))
        drawing.RedrawShadedViews()
        # Alphacam иногда обновляет экстенты SolidBody с задержкой. Повторная
        # коррекция гарантирует привязку дальнейшей нижней левой точки к нулю.
        correction_x = -float(body.MinX)
        correction_y = -float(body.MaxY)
        correction_z = -float(body.MinZ)
        if max(abs(correction_x), abs(correction_y), abs(correction_z)) > 1e-7:
            body.MoveG(correction_x, correction_y, correction_z)
            drawing.RedrawShadedViews()
        self.log.info(
            "ROT после поворота: X[%.6f..%.6f] Y[%.6f..%.6f] Z[%.6f..%.6f]",
            float(body.MinX), float(body.MaxX),
            float(body.MinY), float(body.MaxY),
            float(body.MinZ), float(body.MaxZ),
        )

    def rotate_180_x_and_normalize(self, drawing: Any) -> None:
        """Поворот для -OBR: X+, Y-, Z+ относительно начала координат."""
        body = self.get_solid(drawing)
        center_x = (float(body.MinX) + float(body.MaxX)) / 2.0
        center_y = (float(body.MinY) + float(body.MaxY)) / 2.0
        center_z = (float(body.MinZ) + float(body.MaxZ)) / 2.0
        body.RotateG(
            180.0,
            center_x,
            center_y,
            center_z,
            center_x + 1.0,
            center_y,
            center_z,
        )

        # MinX=0, MaxY=0, MinZ=0 -> X положительный, Y отрицательный, Z положительный.
        body.MoveG(-float(body.MinX), -float(body.MaxY), -float(body.MinZ))
        drawing.RedrawShadedViews()

        correction_x = -float(body.MinX)
        correction_y = -float(body.MaxY)
        correction_z = -float(body.MinZ)
        if max(abs(correction_x), abs(correction_y), abs(correction_z)) > 1e-7:
            body.MoveG(correction_x, correction_y, correction_z)
            drawing.RedrawShadedViews()

        self.log.info(
            "OBR после поворота X: X[%.6f..%.6f] Y[%.6f..%.6f] Z[%.6f..%.6f]",
            float(body.MinX), float(body.MaxX),
            float(body.MinY), float(body.MaxY),
            float(body.MinZ), float(body.MaxZ),
        )

    def export_nc(self, drawing: Any, destination: Path) -> None:
        drawing.OutputNC(str(destination.resolve()), NC_OUTPUT_MODE, True)
        if not destination.is_file():
            raise ProcessingError(f"NC-файл не создан: {destination}")

    def create_rot_program(
        self,
        source_path: Path,
        autostyle_path: Path,
        export_nc: bool,
        require_features: bool = True,
    ) -> StageResult:
        rot_path = source_path.with_name(f"{source_path.stem}-ROT{source_path.suffix}")
        if rot_path.exists():
            raise ProcessingError(f"ROT-файл уже существует: {rot_path}")

        drawing = self.open_drawing(source_path)
        self.save_as(drawing, rot_path)
        drawing = self.session.active_drawing()
        self.rotate_180_z_and_normalize(drawing)
        dimensions = self.solid_body_dimensions(drawing)
        self.set_drawing_parameters(drawing, dimensions)

        side4 = self.extract_workplanes(drawing, target_side=4)
        created = sum(int(item["created"]) for item in side4)
        if require_features and created <= 0:
            raise ProcessingError("После поворота на Side 4 отверстия не найдены.")

        self.apply_autostyle(autostyle_path)
        drawing.SaveAs(str(rot_path.resolve()))

        nc_path = None
        if export_nc:
            nc_path = rot_path.with_suffix(".anc")
            self.export_nc(drawing, nc_path)

        return StageResult(
            stage="ROT",
            message=f"Создан {rot_path.name}",
            details={"ard": str(rot_path), "side4_created": created, "nc": str(nc_path) if nc_path else None},
        )

    def create_obr_program(
        self,
        source_path: Path,
        autostyle_path: Path,
        export_nc: bool,
        require_features: bool = False,
    ) -> StageResult:
        """Создать -OBR, повернуть вокруг X и обработать обратную сторону."""
        obr_path = source_path.with_name(f"{source_path.stem}-OBR{source_path.suffix}")
        if obr_path.exists():
            raise ProcessingError(f"OBR-файл уже существует: {obr_path}")

        drawing = self.open_drawing(source_path)
        self.save_as(drawing, obr_path)
        drawing = self.session.active_drawing()
        self.rotate_180_x_and_normalize(drawing)
        dimensions = self.solid_body_dimensions(drawing)
        self.set_drawing_parameters(drawing, dimensions)

        obr_workplanes = self.extract_workplanes(drawing, exclude_side=2)
        created = sum(int(item["created"]) for item in obr_workplanes)
        if require_features and created <= 0:
            raise ProcessingError("После поворота на X не найдены геометрии для OBR.")

        self.apply_autostyle(autostyle_path)
        drawing.SaveAs(str(obr_path.resolve()))

        nc_path = None
        if export_nc:
            nc_path = obr_path.with_suffix(".anc")
            self.export_nc(drawing, nc_path)

        return StageResult(
            stage="OBR",
            message=f"Создан {obr_path.name}",
            details={"ard": str(obr_path), "side4_created": created, "nc": str(nc_path) if nc_path else None},
        )

    def process_ard(
        self,
        source_path: Path,
        autostyle_path: Path,
        export_nc: bool,
        reverse_side: bool | None = None,
        rotate_program: bool = False,
    ) -> StageResult:
        drawing = self.open_drawing(source_path)
        dimensions = self.solid_dimensions(drawing)
        self.set_drawing_parameters(drawing, dimensions)

        two_programs = bool(rotate_program or reverse_side is True)
        if reverse_side is True:
            self.create_obr_program(
                source_path,
                autostyle_path,
                export_nc,
                require_features=False,
            )
        if rotate_program:
            self.create_rot_program(
                source_path,
                autostyle_path,
                export_nc,
                require_features=False,
            )
        if two_programs:
            drawing = self.open_drawing(source_path)
            dimensions = self.solid_dimensions(drawing)
            self.set_drawing_parameters(drawing, dimensions)

        extracted = self.extract_workplanes(
            drawing,
            exclude_side=2 if two_programs else None,
        )
        side2_created = sum(
            int(item["created"])
            for item in extracted
            if is_side_name(item["name"], 2)
        )
        self.log.info("Side 2 created geometries: %s", side2_created)

        if side2_created and reverse_side is None and not rotate_program:
            return self.create_rot_program(source_path, autostyle_path, export_nc)

        self.apply_autostyle(autostyle_path)
        drawing.SaveAs(str(source_path.resolve()))

        nc_path = None
        if export_nc:
            nc_path = source_path.with_suffix(".anc")
            self.export_nc(drawing, nc_path)

        return StageResult(
            stage="MAIN",
            message=f"Обработан {source_path.name}",
            details={"ard": str(source_path), "nc": str(nc_path) if nc_path else None},
        )


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Alphacam — присадка")
        self.geometry("900x600")
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.ard_var = tk.StringVar()
        self.autostyle_var = tk.StringVar(value=DEFAULT_AUTOSTYLE)
        self.export_var = tk.BooleanVar(value=False)
        self.progid_var = tk.StringVar(value=os.environ.get("ALPHACAM_PROGID", DEFAULT_PROGID))
        self._build_ui()
        self.after(100, self._drain_log)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="ARD-файл или папка:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.ard_var, width=80).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(frame, text="Выбрать ARD", command=self._choose_ard).grid(row=1, column=2, padx=(8, 0))
        ttk.Button(frame, text="Выбрать папку", command=self._choose_folder).grid(row=1, column=3, padx=(8, 0))

        ttk.Label(frame, text="AutoStyle:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.autostyle_var, width=80).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        ttk.Button(frame, text="Выбрать", command=self._choose_autostyle).grid(row=3, column=3, padx=(8, 0))

        ttk.Label(frame, text="COM ProgID:").grid(row=4, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.progid_var, width=40).grid(row=5, column=0, sticky="w", pady=(0, 8))
        ttk.Checkbutton(frame, text="Экспортировать NC", variable=self.export_var).grid(row=5, column=1, sticky="w")

        buttons = ttk.Frame(frame)
        buttons.grid(row=6, column=0, columnspan=4, sticky="w", pady=(4, 10))
        ttk.Button(buttons, text="Проверить и обработать", command=self._start).pack(side=tk.LEFT)
        ttk.Button(buttons, text="X_T → ARD + обработка", command=self._start_xt_batch).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Копировать лог", command=self._copy_log).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Очистить лог", command=self._clear_log).pack(side=tk.LEFT, padx=(8, 0))

        self.log_text = tk.Text(frame, height=25, wrap="word", state=tk.DISABLED)
        self.log_text.grid(row=7, column=0, columnspan=4, sticky="nsew")
        self.log_text.bind("<Control-c>", self._copy_log_event)
        self.log_text.bind("<Control-C>", self._copy_log_event)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=7, column=4, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=0)
        frame.rowconfigure(7, weight=1)

    def _choose_ard(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Alphacam ARD", "*.ard"), ("Все файлы", "*.*")])
        if path:
            self.ard_var.set(path)

    def _choose_folder(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.ard_var.set(path)

    def _choose_autostyle(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Alphacam AutoStyle", "*.ara"), ("Все файлы", "*.*")])
        if path:
            self.autostyle_var.set(path)

    def _clear_log(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _copy_log(self) -> None:
        try:
            text = self.log_text.get("sel.first", "sel.last")
        except tk.TclError:
            text = self.log_text.get("1.0", "end-1c")
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()

    def _copy_log_event(self, _event=None):
        self._copy_log()
        return "break"

    def _start(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Обработка", "Обработка уже выполняется.")
            return
        input_path = Path(self.ard_var.get().strip())
        autostyle_path = Path(self.autostyle_var.get().strip())
        if not input_path.exists():
            messagebox.showerror("Ошибка", "Выберите существующий ARD-файл или папку.")
            return
        plan_path = input_path if input_path.name == "_batch_plan.txt" else None
        if input_path.is_dir():
            candidate = input_path / "_batch_plan.txt"
            if candidate.is_file():
                plan_path = candidate
            else:
                candidate = input_path / "PR" / "_batch_plan.txt"
                if candidate.is_file():
                    plan_path = candidate
        if plan_path is not None:
            self._process_batch_plan(plan_path)
            return
        if input_path.is_dir() and any(
            path.is_file() and path.suffix.lower() == ".x_t"
            for path in input_path.iterdir()
        ):
            self._start_xt_batch(str(input_path))
            return
        if not autostyle_path.is_file():
            messagebox.showerror("Ошибка", "Укажите существующий файл AutoStyle.")
            return

        self.worker = threading.Thread(
            target=self._worker_main,
            args=(input_path, autostyle_path, self.export_var.get(), self.progid_var.get()),
            daemon=True,
        )
        self.worker.start()

    def _start_xt_batch(self, source: str | None = None) -> None:
        """Последовательный импорт X_T с ручной настройкой каждой детали."""
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Обработка", "Сначала дождитесь завершения текущей обработки.")
            return

        if source is None:
            source = filedialog.askdirectory(title="Выберите папку с файлами X_T")
        if not source:
            return
        x_files = sorted(
            [path for path in Path(source).iterdir()
             if path.is_file() and path.suffix.lower() == ".x_t"],
            key=lambda path: path.name.lower(),
        )
        if not x_files:
            messagebox.showinfo("Подготовка X_T", "В выбранной папке нет файлов X_T.")
            return

        default_template = Path(DEFAULT_TEMPLATE)
        if default_template.is_file():
            template = str(default_template)
        else:
            template = filedialog.askopenfilename(
                title="Выберите шаблон ARD",
                filetypes=[("Alphacam template", "*.adt"), ("Все файлы", "*.*")],
                initialdir=str(default_template.parent),
                initialfile=default_template.name,
            )
            if not template:
                return

        output = Path(source) / "PR"
        output.mkdir(exist_ok=True)
        plan_path = output / "_batch_plan.txt"
        logger = logging.getLogger(f"alphacam_gui.xt_batch.{id(self)}")
        logger.setLevel(logging.INFO)
        handler = QueueLogHandler(self.log_queue)
        logger.handlers.clear()
        logger.addHandler(handler)
        session = AlphacamSession(self.progid_var.get().strip() or DEFAULT_PROGID, logger)

        plan_rows: list[str] = []
        try:
            session.connect()
            processor = AlphacamProcessor(session)
            for index, x_path in enumerate(x_files, start=1):
                drawing = processor.import_x_t(x_path, Path(template))
                settings = self._show_orientation_dialog(
                    processor, drawing, x_path.name, index, len(x_files)
                )
                if settings is None:
                    logger.info("Подготовка отменена на %s", x_path.name)
                    break

                ard_path = output / f"{x_path.stem}.ard"
                if settings["skip"]:
                    logger.info("%s: обработка пропущена, ARD не создаётся", x_path.name)
                else:
                    processor.save_as(drawing, ard_path)
                plan_rows.append(
                    f"{ard_path.name}|{int(settings['groove'])}|"
                    f"{int(settings['reverse'])}|{int(settings['rotate_program'])}|"
                    f"{int(settings['skip'])}"
                )
                logger.info(
                    "Подготовлен %s: паз=%s, обратная сторона=%s, "
                    "ROT=%s, пропуск=%s",
                    ard_path.name,
                    "да" if settings["groove"] else "нет",
                    "да" if settings["reverse"] else "нет",
                    "да" if settings["rotate_program"] else "нет",
                    "да" if settings["skip"] else "нет",
                )

            plan_path.write_text(
                "# AutoCAM batch plan v1\n"
                "# file|groove|reverse|rotate_program|skip\n"
                + "\n".join(plan_rows),
                encoding="utf-8",
            )
            logger.info("План сохранён: %s", plan_path)
            messagebox.showinfo(
                "Подготовка X_T",
                f"Подготовлено ARD: {len(plan_rows)} из {len(x_files)}\n\n"
                f"План сохранён: {plan_path}\n\n"
                "Теперь можно запустить пакетную обработку.",
            )
            if messagebox.askyesno(
                "Пакетная обработка",
                "Запустить обработку подготовленных ARD сейчас?",
                parent=self,
            ):
                self._process_batch_plan(plan_path)
        except Exception as exc:
            logger.error("Ошибка подготовки X_T: %s", exc)
            logger.error(traceback.format_exc())
            messagebox.showerror("Подготовка X_T", str(exc))
        finally:
            session.close()

    def _process_batch_plan(self, plan_path: Path) -> None:
        """Вторая фаза: обновить ARD, затем выбрать паз и обработать файл."""
        autostyle_path = Path(self.autostyle_var.get().strip())
        if not autostyle_path.is_file():
            messagebox.showerror("Ошибка", "Укажите существующий файл AutoStyle.")
            return

        logger = logging.getLogger(f"alphacam_gui.plan.{id(self)}")
        logger.setLevel(logging.INFO)
        handler = QueueLogHandler(self.log_queue)
        logger.handlers.clear()
        logger.addHandler(handler)
        session = AlphacamSession(
            self.progid_var.get().strip() or DEFAULT_PROGID, logger
        )
        processed = 0
        try:
            session.connect()
            processor = AlphacamProcessor(session)
            for raw_line in plan_path.read_text(encoding="utf-8").splitlines():
                line = raw_line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split("|")
                if len(parts) < 3:
                    logger.warning("Пропущена некорректная строка плана: %s", line)
                    continue

                ard_path = plan_path.parent / parts[0]
                has_groove = parts[1] == "1"
                reverse_side = parts[2] == "1"
                rotate_program = len(parts) > 3 and parts[3] == "1"
                skip = len(parts) > 4 and parts[4] == "1"
                if skip:
                    logger.info("%s: пропущен по плану, ARD не создавался", parts[0])
                    continue
                if not ard_path.is_file():
                    logger.error("ARD из плана не найден: %s", ard_path)
                    continue

                drawing = processor.open_drawing(ard_path)
                dimensions = processor.solid_dimensions(drawing)
                processor.set_drawing_parameters(drawing, dimensions)

                if has_groove:
                    messagebox.showinfo(
                        "Выбор паза",
                        f"Деталь: {ard_path.name}\n\n"
                        "Параметры чертежа уже обновлены.\n"
                        "Теперь выберите грань/геометрию паза в Alphacam.",
                        parent=self,
                    )
                    processor.capture_groove_via_vba(drawing)
                    drawing.SaveAs(str(ard_path.resolve()))
                else:
                    logger.info("%s: паз не требуется", ard_path.name)

                result = processor.run_stage(
                    f"обработка {ard_path.name}",
                    lambda ard_path=ard_path, reverse_side=reverse_side, rotate_program=rotate_program: processor.process_ard(
                        ard_path,
                        autostyle_path,
                        self.export_var.get(),
                        reverse_side=reverse_side,
                        rotate_program=rotate_program,
                    ),
                )
                logger.info("RESULT %s", result)
                processed += 1

            messagebox.showinfo(
                "Пакетная обработка",
                f"Обработано деталей: {processed}",
                parent=self,
            )
        except Exception as exc:
            logger.error("Ошибка пакетной обработки: %s", exc)
            logger.error(traceback.format_exc())
            messagebox.showerror("Пакетная обработка", str(exc), parent=self)
        finally:
            session.close()

    def _show_orientation_dialog(
        self,
        processor: AlphacamProcessor,
        drawing: Any,
        filename: str,
        index: int,
        total: int,
    ) -> dict[str, Any] | None:
        dialog = tk.Toplevel(self)
        dialog.title(f"Настройка детали {index} из {total}")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        rotation = [0]
        confirmed = [False]
        groove = tk.BooleanVar(value=False)
        reverse = tk.BooleanVar(value=False)
        rotate_program = tk.BooleanVar(value=False)
        skip = tk.BooleanVar(value=False)

        ttk.Label(dialog, text=filename, font=("", 11, "bold")).pack(padx=20, pady=(16, 4))
        ttk.Label(dialog, text="Измените положение детали в Alphacam, затем подтвердите.").pack(padx=20, pady=(0, 12))

        controls = ttk.Frame(dialog)
        controls.pack(padx=20)
        ttk.Button(
            controls, text="Повернуть влево 90°",
            command=lambda: self._try_rotate(processor, drawing, -90, "Z", rotation),
        ).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(
            controls, text="Повернуть вправо 90°",
            command=lambda: self._try_rotate(processor, drawing, 90, "Z", rotation),
        ).grid(row=0, column=1, padx=4, pady=4)
        ttk.Button(
            controls, text="Перевернуть",
            command=lambda: self._try_rotate(processor, drawing, 180, "Y", rotation),
        ).grid(row=1, column=0, columnspan=2, padx=4, pady=4)

        ttk.Checkbutton(
            dialog, text="Паз (необходим выбор геометрии)", variable=groove,
        ).pack(anchor="w", padx=24, pady=(12, 2))
        ttk.Checkbutton(
            dialog,
            text="Обратная сторона (создать стандартную и обратную программы)",
            variable=reverse,
        ).pack(anchor="w", padx=24)
        ttk.Checkbutton(
            dialog,
            text="Поворот (создать стандартную и *-ROT на 180°)",
            variable=rotate_program,
        ).pack(anchor="w", padx=24)
        ttk.Checkbutton(
            dialog, text="Пропускаем обработку (ARD не создаётся)", variable=skip,
        ).pack(anchor="w", padx=24)

        actions = ttk.Frame(dialog)
        actions.pack(pady=(14, 16))

        def accept() -> None:
            confirmed[0] = True
            dialog.destroy()

        ttk.Button(actions, text="Подтвердить и сохранить", command=accept).pack(side=tk.LEFT, padx=4)
        ttk.Button(actions, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=4)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        self.wait_window(dialog)
        if not confirmed[0]:
            return None
        return {
            "rotation": rotation[0] % 360,
            "groove": groove.get(),
            "reverse": reverse.get(),
            "rotate_program": rotate_program.get(),
            "skip": skip.get(),
        }

    def _try_rotate(
        self,
        processor: AlphacamProcessor,
        drawing: Any,
        angle: float,
        axis: str,
        rotation: list[float],
    ) -> None:
        try:
            processor.rotate_imported_solid(drawing, angle, axis)
            rotation[0] += angle
            self.log_queue.put(f"Положение детали изменено: {angle:+g}° вокруг {axis}")
        except Exception as exc:
            messagebox.showerror("Положение детали", str(exc), parent=self)

    def _worker_main(self, input_path: Path, autostyle_path: Path, export_nc: bool, progid: str) -> None:
        logger = logging.getLogger(f"alphacam_gui.{id(threading.current_thread())}")
        logger.setLevel(logging.INFO)
        handler = QueueLogHandler(self.log_queue)
        logger.handlers.clear()
        logger.addHandler(handler)

        session = AlphacamSession(progid, logger)
        try:
            session.connect()
            processor = AlphacamProcessor(session)
            files = sorted(input_path.glob("*.ard")) if input_path.is_dir() else [input_path]
            for file_path in files:
                result = processor.run_stage(
                    f"обработка {file_path.name}",
                    lambda file_path=file_path: processor.process_ard(file_path, autostyle_path, export_nc),
                )
                logger.info("RESULT %s", result)
        except Exception as exc:
            logger.error("FAILED: %s", exc)
            logger.error(traceback.format_exc())
        finally:
            session.close()

    def _drain_log(self) -> None:
        try:
            while True:
                line = self.log_queue.get_nowait()
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, line + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        self.after(100, self._drain_log)


class QueueLogHandler(logging.Handler):
    def __init__(self, target: queue.Queue[str]):
        super().__init__()
        self.target = target

    def emit(self, record: logging.LogRecord) -> None:
        self.target.put(self.format(record))


if __name__ == "__main__":
    app = App()
    app.mainloop()
