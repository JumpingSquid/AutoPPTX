def color_map(index):
    color_map_dict = {
        "sunshine": [[212, 161, 167], [235, 214, 120], [254, 221, 98], [235, 171, 120], [235, 131, 95], [245, 63, 43]],
        "forest": [[87, 151, 115], [145, 96, 38], [166, 164, 167], [229, 233, 239], [89, 56, 0]],
        "sea": [[23, 55, 94], [103, 149, 222], [112, 193, 226], [3, 193, 226], [3, 193, 161]],
        "sea_reverse": [[3, 193, 161], [3, 193, 226], [112, 193, 226], [103, 149, 222], [180, 214, 219]],
        "coldwarm": [[245, 227, 234], [212, 161, 167], [235, 214, 120], [112, 193, 226], [103, 149, 222], [180, 214, 219]]
    }
    return color_map_dict[index]


def textbox(key):
    text_dict = {
        "title_font": "微軟正黑體",
        "comment_font": "微軟正黑體",
        "chart_appendix_font": "微軟正黑體",
        "sample_warn_font": "微軟正黑體",
        "sample_size_font": "微軟正黑體",
        "sample_warn_text": "Sample size < 15",
        "sample_size_text": "總樣本數:"
    }
    return text_dict[key]