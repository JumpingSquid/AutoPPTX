from lxml import etree


class LayoutScanner:

    def __init__(self, filepath=None):
        if filepath is None:
            self.tree = etree.parse(open("C:/Users/user/Desktop/slide1.xml", encoding="utf8"))
        else:
            self.tree = etree.parse(open(filepath, encoding="utf8"))

    def layout_scan(self):
        ns = self.tree.getroot().nsmap
        presult = self.tree.find(".//p:cSld/p:spTree", ns)
        layout_lst = []
        for ele in presult:
            if "graphicFrame" in str(ele):
                frame_name = ele.find(".//p:nvGraphicFramePr/p:cNvPr", ns).attrib["name"]
                aoff = ele.find(".//p:xfrm/a:off", ns).attrib
                aext = ele.find(".//p:xfrm/a:ext", ns).attrib
                layout_lst.append([frame_name, aoff, aext])
            if ("sp" in str(ele)) and (ele.find("p:spPr/", ns) is not None):
                frame_name = ele.find(".//p:nvSpPr/p:cNvPr", ns).attrib["name"]
                aoff = ele.find(".//p:spPr/a:off", ns).attrib
                aext = ele.find(".//p:spPr/a:ext", ns).attrib
                layout_lst.append([frame_name, aoff, aext])

        return layout_lst

    def chart_layout_export(self):
        full_layout = self.layout_scan()
        print(full_layout)
        chart_layout = {"origin": [], "size": []}
        for obj in full_layout:
            if "chart" in obj[0]:
                chart_layout["origin"].append([int(obj[1]["x"]), int(obj[1]["y"])])
                chart_layout["size"].append([int(obj[2]["cx"]), int(obj[2]["cy"])])
        return chart_layout

