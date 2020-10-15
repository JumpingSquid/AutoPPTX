

def order_set():
    sat_item = ['非常不满意', "非常差", "很不满意", "很差", "差",
                '不满意', '有點不满意', "普通", '還可以', '满意', "很满意", '非常满意']
    sat_item.reverse()
    orderset = {
                "satisfaction": ['非常不滿意', "很不滿意", '不滿意', '有點不滿意',
                                 "普通", '還可以', '滿意', "很滿意", '非常滿意'],
                "satisfaction_simplified": sat_item,
                "gender": ['男', "女"],
                "location": ['北', "中", "南", "東", "海外"],
                "nps": ["detractor", "passive", "promoter"],
                "restaurant": ["总和", '云天阁中餐厅', '天星西餐厅（点餐）', '自助餐', "以上皆非"],
                "reserve": ['其他预订方式：', '不是本人预订',
                            '没有预订，现场带位', '线上平台（如: 大众点评、美团等）', '致电餐厅'],
                "brand_preference": ["总和", "国际连锁品牌", "国内连锁品牌", "没有其他偏好"]
                }

    return orderset


def color_map(index):
    color_map_dict = {
        "sunshine": [[212, 161, 167], [235, 214, 120], [254, 221, 98], [235, 171, 120], [235, 131, 95], [245, 63, 43]],
        "forest": [[87, 151, 115], [145, 96, 38], [166, 164, 167], [229, 233, 239], [89, 56, 0]],
        "sea": [[23, 55, 94], [103, 149, 222], [112, 193, 226], [3, 193, 226], [3, 193, 161]],
        "sea_reverse": [[3, 193, 161], [3, 193, 226], [112, 193, 226], [103, 149, 222], [180, 214, 219]],
        "coldwarm": [[245, 227, 234], [212, 161, 167], [235, 214, 120], [112, 193, 226], [103, 149, 222], [180, 214, 219]]
    }
    return color_map_dict[index]


def column_name_map():
    column_name_translate = {
        "meal_sat": "菜品",
        "snack_sat": "點心",
        "drink_sat": "酒水飲料",
        "service_sat": "服務",
        "env_sat": "環境",
        "meeting_room_env": "會議室環境",
        "meeting_room_fac": "會議室設備",
        "meeting_room_drink": "會議室飲品",
        "meeting_room_attitude": "會議室服務"
    }
    return column_name_translate



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