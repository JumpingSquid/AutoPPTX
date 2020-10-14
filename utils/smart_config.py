import pandas as pd
import numpy as np
import os


class PowerpointConfig:

    def __init__(self):
        self.config = {"dataframe": None, "slide": []}
        self.data = None
        self.column_name_map = {}
        self.comment_for_numerical = {0: "mean_max",
                                      1: "mean_min",
                                      2: "percentage_max",
                                      3: "percentage_min"}

        self.comment_for_categorical = {0: "percentage_max",
                                        1: "percentage_min"}

        self.chart = {0: "hist",
                      1: "stackbar",
                      2: "single_category_multiple_bar",
                      3: "multiple_value_column_bar",
                      4: "multiple_dummy_one_column_push"}

    def configwizard(self):
        # ensure data exists or the file type is either csv or excel file first
        filepath = input("Enter the file path, or nothing for using search mode:")
        if filepath != "":
            try:
                if filepath[-3:] == "csv":
                    self.config["dataframe"] = filepath
                    self.data = pd.read_csv(self.config["dataframe"])
                elif filepath[-4:] == "xlsx":
                    self.config["dataframe"] = filepath
                    self.data = pd.read_excel(self.config["dataframe"])
            except:
                raise ValueError("data not found or wrong file format")

        else:
            print("Default: Searching directory under data")
            temp_file_dict = {}
            temp_file_index = 0
            for directory in os.listdir("data"):
                temp_file_dict[temp_file_index] = directory
                temp_file_index += 1
            folder = temp_file_dict[int(input(temp_file_dict))]
            temp_file_dict = {}
            temp_file_index = 0
            for file in os.listdir("data/"+folder):
                if (".csv" in file) or (".xlsx" in file):
                    temp_file_dict[temp_file_index] = file
                    temp_file_index += 1
            file = temp_file_dict[int(input(temp_file_dict))]
            filepath = "data/" + folder + "/" + file
            self.config["dataframe"] = filepath
            self.data = pd.read_csv(self.config["dataframe"])

        # determine the number of slides first
        slide_num = int(input("How many slide do you want to create?"))

        # loop: for each slide ask user to select the categorical  -  value (one variable in this ver.) pair
        # categorical: string, column name, used for segmentation comparison
        # categorical_map: string, column map name, used to provide the correct name for content in slides
        # value: string, column name, can be categorical or numerical variable
        prev_categorical = None
        prev_categorical_map = None
        for page_index in range(0, slide_num):
            print("now we are creating the slide", page_index+1)
            self.config["slide"].append({"page": page_index+1, "data": {}, "comment": None, "charttype": None})

            # create charttype
            # chart type 'multiple_value_column_bar' and 'multiple_dummy_one_column_push' have specific config
            # call corresponding function
            self.config["slide"][page_index]["charttype"] = self.chart[int(input("What is the chart you want to use\n"
                                                                                 "currently support:" + str(self.chart)))]

            if (self.config["slide"][page_index]["charttype"] == "multiple_value_column_bar") | (
                    "multiple_dummy_one_column_push" == self.config["slide"][page_index]["charttype"]):
                self.config["slide"][page_index]["data"]["multicolmap"] = self.column_map_input()

            else:
                # create data - column config
                if prev_categorical is None:
                    prev_categorical = input("Enter the categorical variable you want to use, or input \"auto\" to"
                                             " find possible categorical variable:")
                    if prev_categorical == "auto":
                        prev_categorical = self.autosearch_categorical()
                    prev_value = input("Which value variable you want to use:")
                    assert prev_categorical in self.data.columns
                    assert prev_value in self.data.columns
                    self.config["slide"][page_index]["data"]["columns"] = [prev_categorical, prev_value]
                else:
                    temp_categorical = input("You use: " + prev_categorical + " as category in the last slide,"
                                             " enter nothing for useing the same or enter new category")
                    temp_value = input("You use: " + prev_value + " as value in the last slide,"
                                       " enter nothing for useing the same or enter new category")
                    if temp_categorical == "":
                        self.config["slide"][page_index]["data"]["columns"] = [prev_categorical]
                    elif temp_categorical == "auto":
                        temp_categorical = self.autosearch_categorical()
                        self.config["slide"][page_index]["data"]["columns"] = [temp_categorical]
                        prev_categorical = temp_categorical

                    else:
                        self.config["slide"][page_index]["data"]["columns"] = [temp_categorical]
                        prev_categorical = temp_categorical

                    if temp_value == "":
                        self.config["slide"][page_index]["data"]["columns"].append(prev_value)
                    else:
                        self.config["slide"][page_index]["data"]["columns"].append(temp_value)
                        prev_value = temp_value

                # create data - column_map config
                if prev_categorical_map is None:
                    prev_categorical_map = input("Which categorical variable name you want to use:")
                    prev_value_map = input("Which value variable name you want to use:")
                    self.column_name_map[prev_categorical] = prev_categorical_map
                    self.column_name_map[prev_value] = prev_value_map
                    self.config["slide"][page_index]["data"]["column_name_map"] = [prev_categorical_map, prev_value_map]
                else:
                    if prev_categorical in self.column_name_map.keys():
                        temp_categorical_map = input("You use: " + self.column_name_map[prev_categorical] + " as alias,"
                                                     " enter nothing for using the same or enter new category")
                    else:
                        temp_categorical_map = input("Which categorical variable name you want to use:")
                        self.column_name_map[prev_categorical] = temp_categorical_map

                    if prev_value in self.column_name_map.keys():
                        temp_value_map = input("You use: " + self.column_name_map[prev_value] + " as alias,"
                                               " enter nothing for useing the same or enter new category")
                    else:
                        temp_value_map = input("Which value variable name you want to use:")
                        self.column_name_map[prev_value] = temp_value_map

                    if temp_categorical_map == "":
                        self.config["slide"][page_index]["data"]["column_name_map"] = [prev_categorical_map]
                    else:
                        self.config["slide"][page_index]["data"]["column_name_map"] = [temp_categorical_map]
                        prev_categorical_map = temp_categorical_map

                    if temp_value_map == "":
                        self.config["slide"][page_index]["data"]["column_name_map"].append(prev_value_map)
                    else:
                        self.config["slide"][page_index]["data"]["column_name_map"].append(temp_value_map)
                        prev_value_map = temp_value_map

                # creat comment config
                if np.issubdtype(self.data[prev_value], np.number):
                    comment_config = input("What are comments you want to use\n"
                                           "currently support:"+str(self.comment_for_numerical))
                    self.config["slide"][page_index]["comment"] = [self.comment_for_numerical[int(x)] for x
                                                                   in comment_config.split(",")]
                else:
                    comment_config = input("What are comments you want to use\n"
                                           "currently support"+str(self.comment_for_categorical))
                    self.config["slide"][page_index]["comment"] = [self.comment_for_categorical[int(x)] for x
                                                                   in comment_config.split(",")]

            # create data - title
            self.config["slide"][page_index]["title"] = input("What is the title for the slide?")

            print("slide", page_index+1, "config is created")

        return self.config

    def autosearch_categorical(self):
        categorical_candidate = []
        for column in self.data.columns:
            score = 0
            score += 0.5 * int(column in ["gender", "sex", "location", "work_category", "work_location", "age"])
            score += 0.2 * len(self.data.loc[:, column].dropna())/len(self.data.loc[:, column])
            score += 0.2 * int(len(set(self.data.loc[:, column])) < 10)
            score += max(0.1 - len(set(self.data.loc[:, column]))/len(self.data.loc[:, column]), 0)
            score += (0.1 * int(self.data.loc[:, column].dtype != np.number))
            categorical_candidate.append((column, score))
        categorical_candidate.sort(key=lambda tup: tup[1])
        for ele in categorical_candidate:
            print(ele)
        choice = input("Please select one")
        return choice

    def column_map_input(self):
        # column_agg format:
        # [[CATEGORY_NAME_1, CATEGORY_NAME_2,...], [COL1, COL2, COL3,...], [COL_NAME_1, COL_NAME_2, COL_NAME_3,...]]
        print("column_agg format")
        print("[CATEGORY_NAME_1, CATEGORY_NAME_2,...], [COL1, COL2, COL3,...], [COL_NAME_1, COL_NAME_2, COL_NAME_3,...]")
        multi_column_map = []
        category_list = [x for x in input("Input1: CATEGORY_NAME_1,CATEGORY_NAME_2,...").split(",")]
        column_list = [x for x in input("COLUMN_1,COLUMN_2,...").split(",")]
        column_name_list = [x for x in input("COLUMN_NAME_1,COLUMN_NAME__2,...").split(",")]
        multi_column_map.append({"category_list": category_list,
                                 "column_list": column_list,
                                 "column_name_list": column_name_list})
        return multi_column_map
