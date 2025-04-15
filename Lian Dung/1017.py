from pyautocad import Autocad, APoint
def copy_all_objects(source_file, target_file):
    # 初始化 AutoCAD
    acad = Autocad(create_if_not_exists=True)

    try:
        # 打开来源文件
        source_doc = acad.Application.Documents.Open(source_file)
        if source_doc is None:
            print("未能打开源文件。")
            return
        
        # 打开目标文件
        target_doc = acad.Application.Documents.Open(target_file)
        if target_doc is None:
            print("未能打开目标文件。")
            return

        # 记录源文档中所有对象的 ObjectName
        object_names = set()
        for entity in source_doc.ModelSpace:
            object_names.add(entity.ObjectName)

        print(f"源文档中包含的对象类型: {object_names}")

        # 复制来源文件中的所有元件
        for entity in source_doc.ModelSpace:
            layer_name = entity.Layer  # 获取当前实体的图层名称
            try:
                if entity.ObjectName == 'AcDbLine':
                    # 处理直线
                    if layer_name not in target_doc.Layers:
                        target_doc.Layers.Add(layer_name)
                    start_point = APoint(*entity.StartPoint)
                    end_point = APoint(*entity.EndPoint)
                    new_line = target_doc.ModelSpace.AddLine(start_point, end_point)
                    new_line.Layer = layer_name
                    # print(f"已复制直线: 起点 {start_point}, 终点 {end_point} 到目标文件。")

                elif entity.ObjectName == 'AcDbArc':
                    # 处理弧线
                    if layer_name not in target_doc.Layers:
                        target_doc.Layers.Add(layer_name)
                    center_point = APoint(*entity.Center)
                    radius = entity.Radius
                    start_angle = entity.StartAngle
                    end_angle = entity.EndAngle
                    new_arc = target_doc.ModelSpace.AddArc(center_point, radius, start_angle, end_angle)
                    new_arc.Layer = layer_name
                    # print(f"已复制弧线: 中心点 {center_point}, 半径 {radius} 到目标文件。")

                elif entity.ObjectName == 'AcDbCircle':
                    # 处理圆
                    if layer_name not in target_doc.Layers:
                        target_doc.Layers.Add(layer_name)
                    center_point = APoint(*entity.Center)
                    radius = entity.Radius
                    new_circle = target_doc.ModelSpace.AddCircle(center_point, radius)
                    new_circle.Layer = layer_name
                    # print(f"已复制圆: 中心点 {center_point}, 半径 {radius} 到目标文件。")
                
                elif entity.ObjectName == 'AcDbText':
                    # 处理单行文本
                    new_text = target_doc.ModelSpace.AddText(entity.TextString, APoint(*entity.InsertionPoint), entity.Height)
                    new_text.Layer = layer_name  # 设置图层
                    # print(f"已复制文本: {entity.TextString} 到目标文件。")

                elif entity.ObjectName == 'AcDbMText':
                    # 处理多行文本
                    new_mtext = target_doc.ModelSpace.AddMText(APoint(*entity.InsertionPoint), entity.Width, entity.TextString)
                    new_mtext.Layer = layer_name  # 设置图层
                    # print(f"已复制多行文本: {entity.TextString} 到目标文件。")

        # 处理其他对象类型...


            except Exception as e:
                print(f"处理 {entity.ObjectName} 时发生错误: {e}")

        # 保存并关闭文件
        try:
            target_doc.Save()
        except Exception as e:
            print(f"保存目标文档时发生错误: {e}")

    except Exception as e:
        print(f"操作过程中发生错误: {e}")

    finally:
        # 确保在操作结束后关闭文档
        if 'source_doc' in locals():
            source_doc.Close(False)
        if 'target_doc' in locals():
            target_doc.Close()

    print("所有元件已成功复制。")

# 指定來源和目標檔案的路徑
source_file = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\LT-202-UL.dwg"
target_file = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\成品圖框（中文版) - 複製.dwg"

# 執行複製操作
copy_all_objects(source_file, target_file)
