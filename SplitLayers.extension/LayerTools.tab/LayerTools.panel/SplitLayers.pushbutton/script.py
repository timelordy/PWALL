# -*- coding: utf-8 -*-
"""Разбивка составной стены на части (Parts) стандартными средствами Revit."""

__title__ = 'Разбить\nСтену'
__author__ = 'Wall Layers Separator'

import clr

from System import Enum, MissingMemberException
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


_NAME_PARAMETER_CANDIDATES = []


def _append_name_parameter(param_name):
    """Пытается добавить BuiltInParameter по имени, если он существует."""

    try:
        param = getattr(BuiltInParameter, param_name)
    except (AttributeError, MissingMemberException):
        param = None
    except Exception:
        param = None

    if param is not None:
        _NAME_PARAMETER_CANDIDATES.append(param)


for _param_name in (
    'ALL_MODEL_TYPE_NAME',
    'SYMBOL_NAME_PARAM',
    'ELEM_FAMILY_AND_TYPE_PARAM',
    'PART_MATERIAL_NAME',
):
    _append_name_parameter(_param_name)


def _try_get_enum_value(enum_type, name):
    """Возвращает значение перечисления по имени, если оно существует."""

    if not enum_type or not name:
        return None

    try:
        return getattr(enum_type, name)
    except (AttributeError, MissingMemberException):
        pass
    except Exception:
        pass

    try:
        return Enum.Parse(enum_type, name)
    except Exception:
        return None


def _get_enum_value(enum_type, preferred_names=(), keyword_sets=()):
    """Пытается безопасно найти значение перечисления для разных версий Revit."""

    for name in preferred_names:
        value = _try_get_enum_value(enum_type, name)
        if value is not None:
            return value

    try:
        available_names = list(Enum.GetNames(enum_type))
    except Exception:
        available_names = []

    if not available_names:
        return None

    normalized_names = [(unicode(name), unicode(name).lower()) for name in available_names]

    for keywords in keyword_sets:
        if not keywords:
            continue

        lowered_keywords = [unicode(keyword).lower() for keyword in keywords if keyword]
        if not lowered_keywords:
            continue

        for name, lowered_name in normalized_names:
            if all(keyword in lowered_name for keyword in lowered_keywords):
                value = _try_get_enum_value(enum_type, name)
                if value is not None:
                    return value

    return None


_PARTS_VISIBILITY_SHOW_PARTS = _get_enum_value(
    PartsVisibility,
    preferred_names=('ShowParts', 'ShowPartsOnly'),
    keyword_sets=(('show', 'parts'),),
)

_PARTS_VISIBILITY_SHOW_PARTS_AND_ORIGINAL = _get_enum_value(
    PartsVisibility,
    preferred_names=('ShowPartsAndOriginal', 'ShowOriginalAndParts', 'PartsAndOriginal'),
    keyword_sets=(('show', 'parts', 'original'), ('parts', 'original')),
)


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

    for param_id in _NAME_PARAMETER_CANDIDATES:
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


def _resolve_compound_structure(wall):
    """Пытается извлечь CompoundStructure из стены или её типа."""

    if not wall:
        return None

    candidates = []

    try:
        candidates.append(wall)
    except Exception:
        pass

    try:
        wall_type = wall.WallType
    except (AttributeError, MissingMemberException):
        wall_type = None
    except Exception:
        wall_type = None

    if wall_type and wall_type not in candidates:
        candidates.append(wall_type)

    doc = getattr(wall, 'Document', None)
    if doc:
        try:
            type_element = doc.GetElement(wall.GetTypeId())
        except Exception:
            type_element = None
        if type_element and type_element not in candidates:
            candidates.append(type_element)

    for candidate in candidates:
        if not candidate:
            continue

        for accessor in ('GetCompoundStructure', 'get_CompoundStructure'):
            try:
                method = getattr(candidate, accessor)
            except (AttributeError, MissingMemberException):
                method = None
            except Exception:
                method = None

            if callable(method):
                try:
                    structure = method()
                except Exception:
                    structure = None
                if structure:
                    return structure

        try:
            structure = getattr(candidate, 'CompoundStructure')
        except (AttributeError, MissingMemberException):
            structure = None
        except Exception:
            structure = None

        if structure:
            return structure

    return None


def _get_layer_count(structure):
    """Безопасно определяет количество слоёв в CompoundStructure."""

    if not structure:
        return 0

    try:
        layer_count = int(structure.LayerCount)
        if layer_count > 0:
            return layer_count
    except (AttributeError, MissingMemberException, TypeError, ValueError):
        pass
    except Exception:
        pass

    try:
        layers = structure.GetLayers()
    except (AttributeError, MissingMemberException):
        layers = None
    except Exception:
        layers = None

    if layers is None:
        return 0

    try:
        return len(layers)
    except TypeError:
        count = 0
        try:
            for _ in layers:
                count += 1
        except Exception:
            return 0
        return count
    except Exception:
        return 0


class CompositeWallPartsPipeline(object):
    """Реализует пайплайн: подготовка вида → создание Parts → отчёт."""

    def __init__(self):
        self.doc = doc
        self.uidoc = uidoc
        self.view = uidoc.ActiveView
        self.wall = None
        self.wall_type = None
        self.compound_structure = None
        self.layer_count = 0

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

        allowed_visibility_values = tuple(
            value
            for value in (
                _PARTS_VISIBILITY_SHOW_PARTS,
                _PARTS_VISIBILITY_SHOW_PARTS_AND_ORIGINAL,
            )
            if value is not None
        )

        if allowed_visibility_values and current_visibility in allowed_visibility_values:
            return True

        try:
            current_visibility_name = unicode(current_visibility)
        except Exception:
            current_visibility_name = u""

        if not allowed_visibility_values:
            if current_visibility_name and u"part" in current_visibility_name.lower():
                return True

            output.print_md(
                u"⚠️ API Revit не сообщает доступные значения PartsVisibility."
            )
            output.print_md(
                u"   Включите отображение частей вручную перед запуском скрипта."
            )
            return False

        if _PARTS_VISIBILITY_SHOW_PARTS is None:
            output.print_md(
                u"⚠️ Не удалось получить значение режима 'Показывать части' из API."
            )
            output.print_md(
                u"   Включите отображение Parts вручную и повторите попытку."
            )
            return False

        try:
            view.PartsVisibility = _PARTS_VISIBILITY_SHOW_PARTS
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

        enumerated_layers = []

        try:
            raw_layers = structure.GetLayers()
        except (AttributeError, MissingMemberException):
            raw_layers = None
        except Exception:
            raw_layers = None

        if raw_layers is not None:
            try:
                iterator = iter(raw_layers)
            except TypeError:
                iterator = None
            except Exception:
                iterator = None

            if iterator is not None:
                try:
                    for idx, layer_obj in enumerate(raw_layers):
                        enumerated_layers.append((idx, layer_obj))
                except Exception:
                    enumerated_layers = []

        if not enumerated_layers:
            layer_count = _get_layer_count(structure)
            for index in range(layer_count):
                enumerated_layers.append((index, None))

        for index, layer_obj in enumerated_layers:
            width = None
            function = None
            material_id = None
            is_core = False

            if layer_obj is not None:
                try:
                    width = layer_obj.Width
                except Exception:
                    width = None

                try:
                    function = layer_obj.Function
                except Exception:
                    function = None

                try:
                    material_id = layer_obj.MaterialId
                except Exception:
                    material_id = None

                try:
                    is_core = bool(layer_obj.IsCore)
                except Exception:
                    is_core = False

            if width is None:
                try:
                    width = structure.GetLayerWidth(index)
                except Exception:
                    width = 0.0

            if function is None:
                try:
                    function = structure.GetLayerFunction(index)
                except Exception:
                    function = _DEFAULT_LAYER_FUNCTION

            if material_id is None:
                try:
                    material_id = structure.GetMaterialId(index)
                except Exception:
                    material_id = ElementId.InvalidElementId

            if not is_core:
                try:
                    is_core = structure.IsCoreLayer(index)
                except Exception:
                    is_core = False

            material = self.doc.GetElement(material_id) if material_id else None

            layers.append({
                'index': index + 1,
                'raw_index': index,
                'width': width,
                'width_mm': convert_to_mm(width),
                'material_id': material_id,
                'material_name': get_element_name(material, default=u"Без материала"),
                'function': function,
                'is_core': bool(is_core),
            })

        self.layer_count = max(self.layer_count, len(layers))

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
        self.wall_type = None

        if hasattr(wall, 'WallType'):
            try:
                candidate_type = wall.WallType
            except (AttributeError, MissingMemberException):
                candidate_type = None
            except Exception:
                candidate_type = None

            if isinstance(candidate_type, WallType):
                self.wall_type = candidate_type

        if not isinstance(self.wall_type, WallType):
            try:
                type_element = self.doc.GetElement(wall.GetTypeId())
            except Exception:
                type_element = None

            if isinstance(type_element, WallType):
                self.wall_type = type_element

        self.compound_structure = _resolve_compound_structure(wall)
        self.layer_count = _get_layer_count(self.compound_structure)

        if self.layer_count < 1:
            forms.alert(
                u"Стена не имеет многослойной структуры. Штатные Parts для неё создать нельзя.\n"
                u"Получено количество слоёв: {}. Убедитесь, что выбран тип Basic Wall и у него заданы слои.".format(
                    self.layer_count
                ),
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
        output.print_md(
            u"ℹ️ В типе обнаружено {} слоёв. Ниже приведены подробности по каждому.".format(len(layers))
        )
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
