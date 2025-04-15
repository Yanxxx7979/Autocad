from ezdxf.addons import odafc
import ezdxf
from osgeo import ogr
import sys
import os
import json
import re

class dwg_analysis:
    def __init__(self, filepath, format):
        self.filepath = filepath
        self.format = format

    def dwg2data(self):
        dxf_file = os.path.join(os.path.dirname(self.filepath), os.path.basename(self.filepath).split(".")[0] + ".dxf")
        odafc.convert(self.filepath, dxf_file, version='R2000', replace=True)  
        
        doc = ezdxf.readfile(dxf_file)
        msp = doc.modelspace()

        data = {}
                 # 获取TEXT实体
        texts = msp.query('TEXT')
        text_data = []
        if format == "dwg_txt":
            for text in texts:
                decoded_str = re.sub(r'\\U\+([0-9A-Fa-f]{4})', lambda m: chr(int(m.group(1), 16)), text.dxf.text)
                text_data.append(decoded_str)
            filtered_list = [item for item in text_data if not (isinstance(item, (int, float)) or (isinstance(item, str) and str.isdigit(item)) or (isinstance(item, str) and item.isdigit()))]
            data['TEXT'] = self.remove_duplicates(filtered_list)
            return data

        for text in texts:
            decoded_str = re.sub(r'\\U\+([0-9A-Fa-f]{4})', lambda m: chr(int(m.group(1), 16)), text.dxf.text)
            text_info = {
                'text': decoded_str,
                'insert': (text.dxf.insert[0], text.dxf.insert[1]),
                'height': text.dxf.height,
                'rotation': text.dxf.rotation,
                'style': text.dxf.style,
                'layer': text.dxf.layer
            }
            text_data.append(text_info)

        data['TEXT'] = text_data
                # 获取LINE实体
        lines = msp.query('LINE')
        line_data = []
        for line in lines:
            line_data.append({
                'start': (line.dxf.start[0], line.dxf.start[1]),
                'end': (line.dxf.end[0], line.dxf.end[1])
            })
        data['LINE'] = line_data

        # 获取POLYLINE实体
        polylines = msp.query('POLYLINE')
        polyline_data = []
        for polyline in polylines:
            points = []
            for point in polyline.points():
                points.append((point[0], point[1]))
            polyline_data.append(points)
        data['POLYLINE'] = polyline_data

        # 获取CIRCLE实体
        circles = msp.query('CIRCLE')
        circle_data = []
        for circle in circles:
            circle_data.append({
                'center': (circle.dxf.center[0], circle.dxf.center[1]),
                'radius': circle.dxf.radius
            })
        data['CIRCLE'] = circle_data

        # 获取ARC实体
        arcs = msp.query('ARC')
        arc_data = []
        for arc in arcs:
            arc_data.append({
                'center': (arc.dxf.center[0], arc.dxf.center[1]),
                'radius': arc.dxf.radius,
                'start_angle': arc.dxf.start_angle,
                'end_angle': arc.dxf.end_angle
            })
        data['ARC'] = arc_data

        # 获取ELLIPSE实体
        ellipses = msp.query('ELLIPSE')
        ellipse_data = []
        for ellipse in ellipses:
            ellipse_data.append({
                'center': (ellipse.dxf.center[0], ellipse.dxf.center[1]),
                'major_axis': (ellipse.dxf.major_axis[0], ellipse.dxf.major_axis[1]),
                'ratio': ellipse.dxf.ratio,
                'start_param': ellipse.dxf.start_param,
                'end_param': ellipse.dxf.end_param
            })
        data['ELLIPSE'] = ellipse_data

        return data
    def remove_duplicates(self,lst):
        res = []
        seen = {}
        for i in lst:
            if i not in seen:
                seen[i] = 1
                res.append(i)
        return res
# 示例调用
if __name__ == "__main__":
    # DWG文件路径
    DWG_path = ".\LT-202-UL.dwg"
    format = "dwg_txt"
    dwf = dwg_analysis(DWG_path,format)
    dwf_txt = dwf.dwg2data()
    output_path = os.path.join(os.path.dirname(DWG_path), os.path.basename(DWG_path).split(".")[0] + "_" +format+ ".json")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(dwf_txt, f, ensure_ascii=False, indent=4)



