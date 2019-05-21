def convertRow(row:list, types:list):
    newList = []
    if len(row) == len(types):
        for i in range(0, len(row)):
            convertedValue = convert(row[i], types[i])
            newList.append(convertedValue)
        return newList
    else:
        print("ERROR row count != types count")
        return []

def convert(m_value:str, m_type:str):
    if m_type == 'table':
        return '{'
        pass
    elif m_type == 'boolean':
        if m_value == 1 or m_value == "TRUE":
            return 'true,'
        else:
            return 'false,'
        pass
    elif m_type == 'string':
        return '"{0}",'.format(m_value)
        pass
    elif m_type == 'int':
        return '{:d},'.format(int(m_value))
        pass
    elif m_type == 'float':
        return '{0},'.format(m_value)
        pass
    else:
        pass

def prepare_sheet(sheet):
    descriptions = sheet.row_values(0)
    types = sheet.row_values(1)
    keys = sheet.row_values(2)
    useless_index_list = get_sheet_useless_key(types)
    values = get_sheet_values(sheet)
    remove_useless(useless_index_list, types, keys, values)
    return types, keys, values

def get_sheet_useless_key(types):
    useless_index_list = []
    # 倒序查找
    for i in range(len(types) - 1, -1, -1):
        if types[i] == '_other':
            useless_index_list.append(i)
    return useless_index_list
    pass

def get_sheet_values(sheet):
    # print(sheet.name)
    values = {}
    for i in range(3, sheet.nrows):
        id = i - 3 + 1 # id 从1开始
        values[id] = sheet.row_values(i)
    return values
    pass

def remove_useless(useless_list, types, keys, values):
    for i in range(0, len(useless_list)):
        index = useless_list[i]
        del types[index]
        del keys[index]
        for key in values.keys():
            del values[key][index]
    pass

def found_tabe_end_marks(types, keys):
    mark_list = {}
    overstack_mark_list = []
    for i in range(0, len(keys)):
        m_type = types[i]
        if m_type == 'table':
            mark_count = keys[i].count('#')
            ishave = False
            for j in range(i + 1, len(keys)):
                if keys[j].count('#') == mark_count:
                    mark_list[j - 1] = mark_count
                    ishave = True
                    break
            if not ishave:
                overstack_mark_list.append(mark_count)
    # print(mark_list)
    overstack_mark_list.reverse()
    # print(overstack_mark_list)
    return mark_list, overstack_mark_list

def convert_sheet(types, keys, values):
    mark_list, overstack_mark_list = found_tabe_end_marks(types, keys)

    new_values = {}
    for key in values.keys():
        row = values[key]
        new_row = convertRow(row, types)
        new_values[key] = new_row

    data_str = 'return {\n'
    for id in new_values.keys():
        values = new_values[id]
        value_str = '\t[{0}] = '.format(id) + '{\n'
        data_str = data_str + value_str

        for i in range(0, len(keys)):
            m_key = keys[i]
            m_type = types[i]
            m_value = values[i]
            value_str = '\t\t'
            value_str = value_str + '{0} = {1}\n'.format(m_key, m_value)
            data_str = data_str + value_str
            if i in mark_list.keys():
                tabStr = '\t\t'
                for _ in range(0, mark_list[i]):
                    tabStr = tabStr + '\t'
                data_str = data_str + tabStr + '},\n'
        for j in range(0, len(overstack_mark_list)):
            tabStr = '\t\t'
            mark_count = overstack_mark_list[j]
            for _ in range(0, mark_count):
                tabStr = tabStr + '\t'
            data_str = data_str + tabStr + '},\n'

        data_str = data_str + '\t},\n'
    data_str = data_str + '}'
    data_str = data_str.replace('#', '\t')
    # print(data_str)
    return data_str



