# -*- coding: utf-8 -*-
"""Разбивка составной стены на части (Parts) стандартными средствами Revit."""

__title__ = 'Разбить\nСтену'
__author__ = 'Wall Layers Separator'

import clr

from System import MissingMemberException
from System.Collections.Generic import List

try:
    unicode  # type: ignore[name-defined]
except NameError:  # pragma: no cover - безопасность для Python 3
    unicode = str

# Импорт Revit API
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    MaterialFunctionAssignment,
    PartUtils,
    PartsVisibility,
    UnitTypeId,
    UnitUtils,
    Wall,
    WallType,
)
from Autodesk.Revit.UI.Selection import ObjectType

# Импорт PyRevit
def _load_pyrevit_modules():
    from pyrevit import revit, forms, script
    return revit, forms, script

revit, forms, script = _load_pyrevit_modules()

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


def convert_to_mm(value):
    """Перевод значения из внутренних единиц Revit в миллиметры."""
    if value is None:
        return 0.0

    try:
        return round(UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters), 1)
    except Exception:
        try:
            return round(value * 304.8, 1)
        except Exception:
            return 0.0


def get_element_name(element, default=u"Без имени"):
    """Безопасно получает имя элемента Revit."""
    if not element:
        return default

    try:
        name = element.Name
        if name:
            return name
    except (AttributeError, MissingMemberException):
        pass

    for param_id in (
        BuiltInParameter.ALL_MODEL_TYPE_NAME,
        BuiltInParameter.SYMBOL_NAME_PARAM,
        BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM,
        BuiltInParameter.PART_MATERIAL_NAME,
    ):
        try:
            param = element.get_Parameter(param_id)
            if param and param.HasValue:
                name = param.AsString()
                if name:
                    return name
        except Exception:
            continue

    try:
        return u"ID {}".format(element.Id.IntegerValue)
    except Exception:
        return default


_DEFAULT_LAYER_FUNCTION = getattr(MaterialFunctionAssignment, 'Other', None)

_LAYER_FUNCTION_MAP = {
    MaterialFunctionAssignment.Structure: u"Несущий слой",
    MaterialFunctionAssignment.Substrate: u"Основание",
    MaterialFunctionAssignment.Insulation: u"Утеплитель",
    MaterialFunctionAssignment.Finish1: u"Отделка (наружная)",
    MaterialFunctionAssignment.Finish2: u"Отделка (внутренняя)",
    MaterialFunctionAssignment.Membrane: u"Мембрана",
}

if _DEFAULT_LAYER_FUNCTION not in _LAYER_FUNCTION_MAP:
    _LAYER_FUNCTION_MAP[_DEFAULT_LAYER_FUNCTION] = u"Прочий слой"


def describe_layer_function(layer_function):
    """Текстовое описание функции слоя."""
    return _LAYER_FUNCTION_MAP.get(layer_function, unicode(layer_function))


class CompositeWallPartsPipeline(object):
    """Реализует пайплайн: подготовка вида → создание Parts → отчёт."""

    def __init__(self):
        self.doc = doc
        self.uidoc = uidoc
        self.view = uidoc.ActiveView
        self.wall = None
        self.wall_type = None
        self.compound_structure = None

    # ---------------------------- служебные методы ----------------------------
    def _ensure_view_ready(self):
        """Включает отображение частей на активном виде."""
        view = self.view
        if not view:
            output.print_md(
                u"⚠️ Активный вид не найден. Части будут созданы, но их может не быть видно."
            )
            return False

        try:
            current_visibility = view.PartsVisibility
        except Exception as exc:
            output.print_md(
                u"⚠️ Не удалось прочитать настройку Parts Visibility: {}".format(exc)
            )
            return False

        if current_visibility in (PartsVisibility.ShowParts, PartsVisibility.ShowPartsAndOriginal):
            return True

        try:
            view.PartsVisibility = PartsVisibility.ShowParts
            output.print_md(
                u"ℹ️ Активный вид переключён на режим отображения частей (Parts)."
            )
            return True
        except Exception as exc:
            output.print_md(
                u"⚠️ Не удалось автоматически включить отображение Parts. Причина: {}".format(exc)
            )
            output.print_md(
                u"   Проверьте, не управляет ли параметром шаблон вида, и при необходимости снимите галочку."
            )
            return False

    def _get_existing_parts(self):
        """Возвращает список уже созданных частей для выбранной стены."""
        if not self.wall:
            return []

        wall_id = self.wall.Id

        try:
            part_ids = PartUtils.GetAssociatedParts(self.doc, wall_id)
        except TypeError:
            try:
                part_ids = PartUtils.GetAssociatedParts(self.doc, wall_id, False)
            except Exception:
                part_ids = []
        except Exception:
            part_ids = []

        if not part_ids:
            return []

        return [pid for pid in part_ids]

    def _collect_layers(self):
        """Формирует подробности о слоях из CompoundStructure."""
        layers = []
        structure = self.compound_structure

        if not structure:
            return layers

        for index in range(structure.LayerCount):
            try:
                width = structure.GetLayerWidth(index)
            except Exception:
                width = 0.0

            try:
                function = structure.GetLayerFunction(index)
            except Exception:
                function = _DEFAULT_LAYER_FUNCTION

            try:
                material_id = structure.GetMaterialId(index)
            except Exception:
                material_id = ElementId.InvalidElementId

            material = self.doc.GetElement(material_id) if material_id else None

            layers.append({
                'index': index + 1,
                'raw_index': index,
                'width': width,
                'width_mm': convert_to_mm(width),
                'material_id': material_id,
                'material_name': get_element_name(material, default=u"Без материала"),
                'function': function,
                'is_core': structure.IsCoreLayer(index) if hasattr(structure, 'IsCoreLayer') else False,
            })

        return layers

    def _report_layers(self, layers):
        """Выводит на экран информацию о слоях."""
        if not layers:
            output.print_md(u"⚠️ В структуре стены не найдено слоёв.")
            return

        output.print_md(u"### Слои составной стены:")
        for layer in layers:
            badge = u" (сердцевина)" if layer.get('is_core') else u""
            output.print_md(
                u"- Слой {idx}: **{mat}** — {width} мм, {func}{badge}".format(
                    idx=layer['index'],
                    mat=layer['material_name'],
                    width=layer['width_mm'],
                    func=describe_layer_function(layer['function']),
                    badge=badge,
                )
            )

    def _report_parts(self, part_ids, layers):
        """Опубликовать список полученных частей."""
        if not part_ids:
            output.print_md(u"⚠️ Для стены не найдено частей.")
            return

        output.print_md(u"### Созданные части:")

        if layers and len(layers) == len(part_ids):
            for layer, part_id in zip(layers, part_ids):
                part = self.doc.GetElement(part_id)
                part_name = get_element_name(part, default=u"Часть")
                output.print_md(
                    u"- Слой {idx} → часть `{name}` (ID {pid}, {width} мм)".format(
                        idx=layer['index'],
                        name=part_name,
                        pid=part_id.IntegerValue,
                        width=layer['width_mm'],
                    )
                )
        else:
            for index, part_id in enumerate(part_ids, 1):
                part = self.doc.GetElement(part_id)
                part_name = get_element_name(part, default=u"Часть")
                output.print_md(
                    u"- Часть {idx}: `{name}` (ID {pid})".format(
                        idx=index,
                        name=part_name,
                        pid=part_id.IntegerValue,
                    )
                )

        output.print_md(u"### Что дальше")
        output.print_md(u"- При необходимости используйте инструмент **Divide Parts**, чтобы дополнительно нарезать части по эскизам или уровням.")
        output.print_md(u"- Для отображения исходной стены вместе с частями переключите вид на режим *Show Parts and Original*.")
        output.print_md(u"- Части находятся в отдельной категории, поэтому для спецификаций и фильтров используйте таблицы Parts.")

    # ------------------------------- рабочий процесс -------------------------------
    def select_wall(self):
        """Запрашивает у пользователя выбор многослойной стены."""
        try:
            reference = self.uidoc.Selection.PickObject(
                ObjectType.Element,
                u"Выберите составную стену для создания частей"
            )
        except Exception as exc:
            forms.alert(u"Стена не выбрана. Причина: {}".format(exc), exitscript=True)
            return False

        wall = self.doc.GetElement(reference.ElementId)
        if not isinstance(wall, Wall):
            forms.alert(u"Выбранный элемент не является стеной.", exitscript=True)
            return False

        self.wall = wall
        self.wall_type = wall.WallType if hasattr(wall, 'WallType') else None

        if isinstance(self.wall_type, WallType):
            try:
                self.compound_structure = self.wall_type.GetCompoundStructure()
            except Exception:
                self.compound_structure = None
        else:
            self.compound_structure = None

        if not self.compound_structure or self.compound_structure.LayerCount < 2:
            forms.alert(
                u"Стена не имеет многослойной структуры. Штатные Parts для неё создать нельзя.",
                exitscript=True
            )
            return False

        return True

    def _create_parts(self):
        """Создаёт части для выбранной стены."""
        if not self.wall:
            return []

        wall_id = self.wall.Id

        try:
            if not PartUtils.IsElementValidForCreateParts(self.doc, wall_id):
                forms.alert(
                    u"Элемент нельзя превратить в Parts. Проверьте, не является ли стена вложенной в группу или ссылку.",
                    exitscript=True
                )
                return []
        except Exception:
            pass

        element_ids = List[ElementId]()
        element_ids.Add(wall_id)

        PartUtils.CreateParts(self.doc, element_ids)
        return self._get_existing_parts()

    def execute(self):
        """Запускает пайплайн."""
        if not self.select_wall():
            return

        wall_name = get_element_name(self.wall_type, default=u"Стена")
        output.print_md(u"## Разбивка стены **{}** на Parts".format(wall_name))

        layers = self._collect_layers()
        self._report_layers(layers)

        existing_parts = self._get_existing_parts()
        if existing_parts:
            output.print_md(
                u"⚠️ Для стены уже созданы части ({} шт.). Скрипт переключит вид и покажет информацию.".format(
                    len(existing_parts)
                )
            )

            with revit.Transaction(u"Актуализация вида (Parts)"):
                self._ensure_view_ready()

            self._report_parts(existing_parts, layers)
            return

        if not forms.alert(
                u"Создать части (Parts) для всех {} слоёв?".format(len(layers)),
                yes=True,
                no=True
        ):
            output.print_md(u"⚠️ Операция отменена пользователем.")
            return

        with revit.Transaction(u"Создание Parts из стены"):
            self._ensure_view_ready()
            part_ids = self._create_parts()

        if not part_ids:
            output.print_md(u"⚠️ Части не были созданы. Проверьте журнал выше.")
            return

        output.print_md(u"## ✅ Стена преобразована в Parts")
        self._report_parts(part_ids, layers)


if __name__ == '__main__':
    pipeline = CompositeWallPartsPipeline()
    pipeline.execute()
