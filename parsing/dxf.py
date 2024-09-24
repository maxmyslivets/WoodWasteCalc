import ezdxf


def dxf_parse(dxf_filename: str) -> tuple:

    doc = ezdxf.readfile(dxf_filename)

    mt_list = [[mt.plain_text().replace('\n', ' '), mt.dxf.insert[1], mt.dxf.insert[0]] for mt in
               doc.modelspace().query('MTEXT')]
    mt_list_sorted_by_x = sorted(mt_list, key=lambda m: m[1])[::-1]
    name_group = mt_list_sorted_by_x.pop(0)[0]

    n_coloumns = 5
    data_array = [mt_list_sorted_by_x[i:i + n_coloumns] for i in range(0, len(mt_list_sorted_by_x), n_coloumns)]
    for idx, row in enumerate(data_array):
        data_array[idx] = sorted(data_array[idx], key=lambda m: m[2])
    table_array = []
    for row in data_array:
        table_array.append([mt[0] for mt in row])

    return table_array, name_group
