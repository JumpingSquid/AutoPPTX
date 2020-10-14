# This module is used to create comment from data
# The requirement of the data is to have only two columns
# The first column should be the categorical
# The second column should be numerical
from autoppt_util.pptx_params import column_name_map, comment_type_language


class CommentCreator:

    def __init__(self, data, comment_config):
        self.data = data
        self.comment_config = comment_config
        self.value_column_translate = column_name_map()

    def comment_create(self):
        if self.comment_config["comment_type"] in ["percentage_max", "%max"]:
            category_count = self.data.loc[:, self.comment_config["cat_col"]].value_counts(normalize=True)
            category_name = category_count.index[0]
            category_value = category_count.iloc[0]
            new_comment = comment_type_language("zh", self.comment_config["comment_type"]).\
                format(self.comment_config["col_alias"], category_name, category_value)

        elif self.comment_config["comment_type"] in ["percentage_min", "%min"]:
            category_count = self.data.loc[:, self.comment_config["cat_col"]].value_counts(normalize=True)
            category_name = category_count.index[-1]
            category_value = category_count.iloc[-1]
            new_comment = comment_type_language("zh", self.comment_config["comment_type"]).\
                format(self.comment_config["col_alias"], category_name, category_value)

        elif self.comment_config["comment_type"] in ["nss_max"]:
            def net_sat_appendix(data, sat_col):
                try:
                    net_sat = (sum(data.loc[:, sat_col].isin(["很满意"])) -
                               sum(data.loc[:, sat_col].isin(["普通", "差", "很差"]))) / \
                              sum(data.loc[:, sat_col].isin(["很满意", "满意", "普通", "差", "很差"]))
                    net_sat = int(net_sat * 100)
                except ZeroDivisionError:
                    net_sat = 0
                return net_sat
            category_name = ""
            category_value = 0
            for value_column in self.comment_config["cat_col"]:
                temp_value = net_sat_appendix(self.data, value_column)
                if temp_value > category_value:
                    category_name = value_column
                    category_value = temp_value

            if category_name in self.value_column_translate.keys():
                category_name = self.value_column_translate[category_name]

            new_comment = comment_type_language("zh", self.comment_config["comment_type"]). \
                format(self.comment_config["col_alias"][category_name], category_value)

        elif self.comment_config["comment_type"] in ["multi_value_percentage_max", "mv%max"]:
            category_name = ""
            category_value = 0
            for value_column in self.comment_config["val_col"]:
                temp_value = sum(self.data[value_column] == self.comment_config["pos_indicator"]) /\
                             len(self.data[value_column])
                if temp_value > category_value:
                    category_name = value_column
                    category_value = temp_value

            if category_name in self.value_column_translate.keys():
                category_name = self.value_column_translate[category_name]

            new_comment = comment_type_language("zh", self.comment_config["comment_type"]).\
                format(self.comment_config["col_alias"], self.comment_config["pos_indicator"], category_name, category_value)

        elif self.comment_config["comment_type"] in ["multi_value_percentage_max", "mv%min"]:
            category_name = ""
            category_value = 1
            for value_column in self.comment_config["val_col"]:
                temp_value = sum(self.data[value_column] == self.comment_config["pos_indicator"]) /\
                             len(self.data[value_column])
                if temp_value < category_value:
                    category_name = value_column
                    category_value = temp_value

            if category_name in self.value_column_translate.keys():
                category_name = self.value_column_translate[category_name]

            new_comment = comment_type_language("zh", self.comment_config["comment_type"]).\
                format(self.comment_config["col_alias"], self.comment_config["pos_indicator"], category_name, category_value)

        else:
            print("WARN: Unknown comment type received, empty comment produced")
            new_comment = ""

        return new_comment
