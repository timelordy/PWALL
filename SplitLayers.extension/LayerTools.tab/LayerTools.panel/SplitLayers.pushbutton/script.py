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
    WallUtils,
)
from Autodesk.Revit.UI.Selection import ObjectType

try:
    from Autodesk.Revit.Exceptions import ArgumentException as RevitArgumentException
except ImportError:
    RevitArgumentException = Exception

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


def _coerce_element_id_list(raw_ids):
    """Преобразует коллекцию ElementId в список с защитой от ошибок."""

    result = []

    if raw_ids is None:
        return result

    try:
        iterator = iter(raw_ids)
    except TypeError:
        iterator = None
    except Exception:
        iterator = None

    if iterator is not None:
        try:
            for item in raw_ids:
                if isinstance(item, ElementId):
                    result.append(item)
                else:
                    try:
                        result.append(ElementId(int(item)))
                    except Exception:
                        continue
        except Exception:
            result = []

        return result

    count = 0

    try:
        count = int(getattr(raw_ids, 'Count', 0))
    except Exception:
        count = 0

    for index in range(count):
        try:
            item = raw_ids[index]
        except Exception:
            continue

        if isinstance(item, ElementId):
            result.append(item)
        else:
            try:
                result.append(ElementId(int(item)))
            except Exception:
                continue

    return result


def _invoke_with_optional_bool(method):
    """Безопасно вызывает метод без аргументов или с булевым аргументом."""

    if not callable(method):
        return None

    try:
        return method()
    except TypeError:
        pass
    except Exception:
        return None

    for flag in (False, True):
        try:
            return method(flag)
        except TypeError:
            continue
        except Exception:
            continue

    return None


def _get_wall_type_kind(wall_type):
    """Пытается определить тип стены (Basic, Stacked и т.д.)."""

    if not isinstance(wall_type, WallType):
        return None

    for accessor in ('Kind', 'get_Kind'):
        try:
            value = getattr(wall_type, accessor)
        except (AttributeError, MissingMemberException):
            value = None
        except Exception:
            value = None

        if value is None:
            continue

        if callable(value):
            try:
                value = value()
            except Exception:
                value = None

        if value is not None:
            return value

    return None


def _is_stacked_wall_kind(kind):
    """Определяет, относится ли стена к семейству "Составная" (Stacked)."""

    if not kind:
        return False

    try:
        kind_text = unicode(kind)
    except Exception:
        try:
            kind_text = str(kind)
        except Exception:
            kind_text = u''

    kind_text = kind_text.lower()
    return 'stacked' in kind_text or u'состав' in kind_text


def _get_stacked_wall_member_ids(wall, wall_type=None):
    """Возвращает ElementId вложенных стен для составной стены."""

    member_ids = []
    seen_ids = set()

    def _extend(result):
        for element_id in _coerce_element_id_list(result):
            try:
                key = element_id.IntegerValue
            except Exception:
                key = None

            if key is not None:
                if key in seen_ids:
                    continue
                seen_ids.add(key)

            member_ids.append(element_id)

    candidates = []

    if wall:
        candidates.append(wall)

    if wall_type and wall_type not in candidates:
        candidates.append(wall_type)

    try:
        utils_method = getattr(WallUtils, 'GetStackedWallMemberIds')
    except Exception:
        utils_method = None

    if callable(utils_method) and wall:
        try:
            _extend(utils_method(wall))
        except Exception:
            pass

    for candidate in candidates:
        for accessor in (
            'GetStackedWallMemberIds',
            'get_StackedWallMemberIds',
            'GetMemberIds',
            'get_MemberIds',
            'GetStackedWallMembers',
            'get_StackedWallMembers',
        ):
            try:
                method = getattr(candidate, accessor)
            except (AttributeError, MissingMemberException):
                method = None
            except Exception:
                method = None

            if not callable(method):
                continue

            try:
                result = _invoke_with_optional_bool(method)
            except Exception:
                result = None

            if result is None:
                continue

            _extend(result)

    return member_ids


def _enumerate_structure_layers(structure):
    """Возвращает пары (индекс, слой) для CompoundStructure."""

    enumerated_layers = []

    if not structure:
        return enumerated_layers

    try:
        raw_layers = structure.GetLayers()
    except (AttributeError, MissingMemberException):
        raw_layers = None
    except Exception:
        raw_layers = None

    if raw_layers is not None:
        try:
            for idx, layer_obj in enumerate(raw_layers):
                enumerated_layers.append((idx, layer_obj))
        except Exception:
            enumerated_layers = []

    if not enumerated_layers:
        layer_count = _get_layer_count(structure)
        for index in range(layer_count):
            enumerated_layers.append((index, None))

    return enumerated_layers


def _structure_to_layer_info(structure, doc, index_offset=0):
    """Преобразует CompoundStructure в список словарей со сведениями о слоях."""

    layers = []

    if not structure:
        return layers

    for index, layer_obj in _enumerate_structure_layers(structure):
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

        material = doc.GetElement(material_id) if material_id else None

        layers.append({
            'index': index_offset + index + 1,
            'raw_index': index,
            'width': width,
            'width_mm': convert_to_mm(width),
            'material_id': material_id,
            'material_name': get_element_name(material, default=u"Без материала"),
            'function': function,
            'is_core': bool(is_core),
        })

    return layers


def _format_layer_part_name(layer):
    """Формирует подпись слоя для использования как имя части."""
    if not layer:
        return None

    try:
        index = int(layer.get('index')) if layer.get('index') is not None else None
    except Exception:
        index = None

    name_bits = []
    if index:
        name_bits.append(u"Слой {}".format(index))

    segment_index = layer.get('segment_index')
    segment_layer_index = layer.get('segment_layer_index')
    segment_name = layer.get('segment_name')
    segment_chunks = []
    if segment_index:
        segment_chunks.append(u"Сегмент {}".format(segment_index))
    if segment_layer_index:
        segment_chunks.append(u"Слой внутри сегмента {}".format(segment_layer_index))
    if segment_name:
        segment_chunks.append(segment_name)
    if segment_chunks:
        name_bits.append(u" / ".join(segment_chunks))

    if not name_bits:
        name_bits.append(u"Слой")

    label = u" / ".join(name_bits)

    description_parts = []
    material_name = layer.get('material_name')
    if material_name:
        description_parts.append(material_name)

    width_mm = layer.get('width_mm')
    if width_mm is not None:
        try:
            description_parts.append(u"{:.1f} мм".format(float(width_mm)))
        except Exception:
            pass

    func_label = describe_layer_function(layer.get('function'))
    if func_label:
        description_parts.append(func_label)

    if layer.get('is_core'):
        description_parts.append(u"Сердцевина")

    if description_parts:
        return u"{} - {}".format(label, u", ".join(description_parts))
    return label


def _set_part_name(part, value):
    """Записывает строку в параметр имени части."""
    if not part or value in (None, ''):
        return False

    try:
        text_value = unicode(value)
    except Exception:
        try:
            text_value = str(value)
        except Exception:
            return False

    param = None
    try:
        param = part.get_Parameter(BuiltInParameter.DPART_PART_NAME)
    except Exception:
        param = None

    if param and not getattr(param, 'IsReadOnly', True):
        try:
            param.Set(text_value)
            return True
        except Exception:
            pass

    try:
        part.Name = text_value
        return True
    except Exception:
        return False


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
        self.wall_kind = None
        self.uses_stacked_members = False
        self.stacked_member_ids = []

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
        structure = self.compound_structure

        layers = _structure_to_layer_info(structure, self.doc, index_offset=0)

        if not layers and self.stacked_member_ids:
            layers = self._collect_layers_from_stacked()

        self.layer_count = max(self.layer_count, len(layers))

        return layers

    def _collect_layers_from_stacked(self):
        """Формирует подробности о слоях для составной (stacked) стены."""

        member_ids = list(self.stacked_member_ids or [])

        if not member_ids:
            member_ids = _get_stacked_wall_member_ids(self.wall, self.wall_type)
            self.stacked_member_ids = member_ids

        if not member_ids:
            return []

        layers = []
        index_offset = 0
        doc = self.doc

        for segment_index, member_id in enumerate(member_ids, 1):
            try:
                member_element = doc.GetElement(member_id)
            except Exception:
                member_element = None

            member_wall = member_element if isinstance(member_element, Wall) else None
            member_type = None

            if isinstance(member_wall, Wall):
                try:
                    candidate_type = member_wall.WallType
                except (AttributeError, MissingMemberException):
                    candidate_type = None
                except Exception:
                    candidate_type = None

                if isinstance(candidate_type, WallType):
                    member_type = candidate_type

            if not isinstance(member_type, WallType):
                if isinstance(member_element, WallType):
                    member_type = member_element
                else:
                    try:
                        candidate_type = doc.GetElement(member_id)
                    except Exception:
                        candidate_type = None

                    if isinstance(candidate_type, WallType):
                        member_type = candidate_type

            if not isinstance(member_type, WallType):
                continue

            structure = _resolve_compound_structure(member_type)
            member_layers = _structure_to_layer_info(structure, doc, index_offset=index_offset)

            if not member_layers:
                continue

            segment_name = get_element_name(member_type, default=u"Сегмент стены")
            segment_offset = index_offset

            for layer in member_layers:
                layer['segment_index'] = segment_index
                layer['segment_name'] = segment_name
                layer['segment_layer_index'] = layer['index'] - segment_offset

            layers.extend(member_layers)
            index_offset += len(member_layers)

        return layers


    def _rename_parts(self, part_ids, layers):
        """Присваивает частям имена, основанные на данных слоёв."""
        if not part_ids:
            return 0

        if not layers:
            iterable = ((part_id, None) for part_id in part_ids)
        elif len(layers) == len(part_ids):
            iterable = zip(part_ids, layers)
        else:
            iterable = zip(part_ids, layers)

        renamed = 0
        for part_id, layer in iterable:
            try:
                part = self.doc.GetElement(part_id)
            except Exception:
                part = None

            if part is None:
                continue

            display_name = _format_layer_part_name(layer) if layer else None
            if not display_name:
                continue

            if _set_part_name(part, display_name):
                renamed += 1

        return renamed


    def _report_layers(self, layers):
        """Выводит на экран информацию о слоях."""
        if not layers:
            output.print_md(u"⚠️ В структуре стены не найдено слоёв.")
            return

        output.print_md(u"### Слои составной стены:")
        for layer in layers:
            badge = u" (сердцевина)" if layer.get('is_core') else u""
            segment_info = u""

            segment_index = layer.get('segment_index')
            if segment_index:
                segment_parts = [u"сегмент {}".format(segment_index)]

                segment_layer_index = layer.get('segment_layer_index')
                if segment_layer_index:
                    segment_parts.append(u"слой {}".format(segment_layer_index))

                segment_name = layer.get('segment_name')
                if segment_name:
                    segment_parts.append(segment_name)

                segment_info = u" [{}]".format(u", ".join(segment_parts))

            output.print_md(
                u"- Слой {idx}: **{mat}** — {width} мм, {func}{badge}{segment_info}".format(
                    idx=layer['index'],
                    mat=layer['material_name'],
                    width=layer['width_mm'],
                    func=describe_layer_function(layer['function']),
                    badge=badge, segment_info=segment_info,
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
                segment_info = u""

                segment_index = layer.get('segment_index')
                if segment_index:
                    segment_parts = [u"сегмент {}".format(segment_index)]

                    segment_layer_index = layer.get('segment_layer_index')
                    if segment_layer_index:
                        segment_parts.append(u"слой {}".format(segment_layer_index))

                    segment_name = layer.get('segment_name')
                    if segment_name:
                        segment_parts.append(segment_name)

                    segment_info = u" [{}]".format(u", ".join(segment_parts))

                output.print_md(
                    u"- Слой {idx}{segment_info} → часть `{name}` (ID {pid}, {width} мм)".format(
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

        self.wall_kind = _get_wall_type_kind(self.wall_type)
        self.compound_structure = _resolve_compound_structure(wall)
        self.layer_count = _get_layer_count(self.compound_structure)
        self.stacked_member_ids = _get_stacked_wall_member_ids(self.wall, self.wall_type)
        self.uses_stacked_members = bool(self.stacked_member_ids) and (
            self.layer_count < 1 or _is_stacked_wall_kind(self.wall_kind)
        )

        if self.layer_count < 1 and not self.stacked_member_ids:
            forms.alert(
                u"Стена не имеет многослойной структуры. Штатные Parts для неё создать нельзя.\n"
                u"Получено количество слоёв: {}. Убедитесь, что выбран тип стены со структурой слоёв или разложите сложные стены на составные сегменты.".format(
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

        try:
            PartUtils.CreateParts(self.doc, element_ids)
        except RevitArgumentException as exc:
            output.print_md(
                u"[!] Не удалось создать Parts для стены (ID {}). Причина: {}".format(
                    wall_id.IntegerValue, exc
                )
            )
            return []
        except Exception as exc:
            output.print_md(
                u"[!] Неожиданная ошибка при создании Parts для стены (ID {}): {}".format(
                    wall_id.IntegerValue, exc
                )
            )
            return []

        part_ids = self._get_existing_parts()

        return part_ids

    def execute(self):
        """Запускает пайплайн."""
        if not self.select_wall():
            return

        wall_name = get_element_name(self.wall_type, default=u"Стена")
        output.print_md(u"## Разбивка стены **{}** на Parts".format(wall_name))

        if self.uses_stacked_members and self.stacked_member_ids:
            output.print_md(
                u"ℹ️ Обнаружена составная стена (Stacked). Будут обработаны {} сегментов.".format(
                    len(self.stacked_member_ids)
                )
            )

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

            with revit.Transaction(u"Обновление частей (Parts)"):
                self._ensure_view_ready()
                renamed = self._rename_parts(existing_parts, layers)

            if renamed:
                output.print_md(
                    u"ℹ️ Обновлены имена {} частей в соответствии со слоями.".format(renamed)
                )

            self._report_parts(existing_parts, layers)
            return

        if not forms.alert(
                u"Создать части (Parts) для всех {} слоёв?".format(len(layers)),
                yes=True,
                no=True
        ):
            output.print_md(u"⚠️ Операция отменена пользователем.")
            return

        if _is_stacked_wall_kind(self.wall_kind):
            wall_name = get_element_name(self.wall, default=u"Стена")
            output.print_md(
                u"[!] Стена **{}** остаётся составной. Автоматическое создание Parts пропускаю.".format(wall_name)
            )
            return

        with revit.Transaction(u"Создание Parts из стены"):
            self._ensure_view_ready()
            part_ids = self._create_parts()

        if not part_ids:
            output.print_md(u"⚠️ Части не были созданы. Проверьте журнал выше.")
            return

        with revit.Transaction(u"Переименование Parts по слоям"):
            renamed = self._rename_parts(part_ids, layers)

        if renamed:
            output.print_md(
                u"ℹ️ Частям присвоены имена по слоям ({} элементов).".format(renamed)
            )

        output.print_md(u"## ✅ Стена преобразована в Parts")
        self._report_parts(part_ids, layers)


if __name__ == '__main__':
    pipeline = CompositeWallPartsPipeline()
    pipeline.execute()
