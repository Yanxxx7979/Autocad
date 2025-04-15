from pyautocad import Autocad, APoint

def merge_drawings_with_specific_areas():
    # 初始化 AutoCAD
    acad = Autocad(create_if_not_exists=True)

    # 打開圖面文件
    path_a = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\LT-202-UL.dwg"
    path_b = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\LT-501-UL.dwg"
    path_c = r"C:\Users\user\Downloads\Lian Dung\c.dwg"
    path_d = r"C:\Users\user\Downloads\Lian Dung\d.dwg"
    target_path = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\成品圖框.dwg"

    try:
        acad_doc_a = acad.Application.Documents.Open(path_a)
        acad_doc_b = acad.Application.Documents.Open(path_b)
        acad_doc_c = acad.Application.Documents.Open(path_c)
        acad_doc_d = acad.Application.Documents.Open(path_d)
        acad_target_doc = acad.Application.Documents.Open(target_path)

        def copy_selected_objects(source_doc, target_doc, min_point, max_point, selection_set_name):
            try:
                # 如果選擇集已存在，先刪除
                for selection in source_doc.SelectionSets:
                    if selection.Name == selection_set_name:
                        selection.Delete()

                selection_set = source_doc.SelectionSets.Add(selection_set_name)

                # 將定義的 min_point, max_point 轉成方便比較的數值
                # 假設您確定 min_point.x < max_point.x，min_point.y < max_point.y
                min_x, max_x = min_point.x, max_point.x
                min_y, max_y = min_point.y, max_point.y

                for entity in source_doc.ModelSpace:
                    try:
                        obj_name = entity.ObjectName  # 例如 'AcDbLine', 'AcDbArc', 'AcDbCircle', 'AcDbText' 等

                        # 先針對不同類型做範圍判斷
                        # ---------------------------------------------------------------------
                        # 1) 文字、塊參考等具有 InsertionPoint 的物件
                        # ---------------------------------------------------------------------
                        if obj_name in ['AcDbText', 'AcDbMText', 'AcDbBlockReference']:
                            insertion_point = APoint(*entity.InsertionPoint)
                            if not (min_x <= insertion_point.x <= max_x and max_y <= insertion_point.y <= min_y):
                                continue
            # min_point_a = APoint(100, 150)
            # max_point_a = APoint(189.8563, 110)
                        # ---------------------------------------------------------------------
                        # 2) 直線
                        # ---------------------------------------------------------------------
                        elif obj_name == 'AcDbLine':
                            start_point = APoint(*entity.StartPoint)
                            end_point = APoint(*entity.EndPoint)
                            # 這邊的範圍判斷可以是「任一端點落在區域內就算」，也可以加更嚴謹的判斷
                            # 以下示範只要 StartPoint 或 EndPoint 位於區域內，就通過
                            if not (min_x <= start_point.x <= max_x and max_y <= start_point.y <= min_y) \
                            and not (min_x <= end_point.x <= max_x and max_y <= end_point.y <= min_y):
                                continue

                        # ---------------------------------------------------------------------
                        # 3) 圓、弧
                        # ---------------------------------------------------------------------
                        elif obj_name in ['AcDbCircle', 'AcDbArc']:
                            center_point = APoint(*entity.Center)
                            
                            if not (min_x <= center_point.x <= max_x and max_y <= center_point.y <= min_y):
                                
                                continue

                        else:
                            # 如果不是上面列出的類型，直接跳過或自行視需求擴充
                            continue

                        # ---------------------------------------------------------------------
                        # 走到這裡表示範圍檢查合格，才複製到目標圖面
                        # 先確保 layer 存在於目標圖面
                        layer_name = entity.Layer
                        if layer_name not in target_doc.Layers:
                            target_doc.Layers.Add(layer_name)

                        # 根據物件類型來建立目標圖面的新物件
                        if obj_name == 'AcDbLine':
                            start_point = APoint(*entity.StartPoint)
                            end_point = APoint(*entity.EndPoint)
                            # print(start_point.x, end_point.x)
                            new_line = target_doc.ModelSpace.AddLine(start_point, end_point)
                            new_line.Layer = layer_name
       
                        elif obj_name == 'AcDbCircle':
                            center_point = APoint(*entity.Center)
                            radius = entity.Radius
                            new_circle = target_doc.ModelSpace.AddCircle(center_point, radius)
                            new_circle.Layer = layer_name

                        elif obj_name == 'AcDbArc':
                            center_point = APoint(*entity.Center)
                            radius = entity.Radius
                            start_angle = entity.StartAngle
                            end_angle = entity.EndAngle
                            new_arc = target_doc.ModelSpace.AddArc(center_point, radius, start_angle, end_angle)
                            new_arc.Layer = layer_name

                        elif obj_name in ['AcDbText', 'AcDbMText']:
                            insertion_point = APoint(*entity.InsertionPoint)
                            # 對於 MText 需要特別方法，簡化只示範 Text
                            if obj_name == 'AcDbText':
                                new_text = target_doc.ModelSpace.AddText(entity.TextString,
                                                                        insertion_point,
                                                                        entity.Height)
                                new_text.Layer = layer_name
                            elif obj_name == 'AcDbMText':
                                # MText 的拷貝，需要用 AddMText，並設定寬度和高度
                                pass

                        # 如果還想拷貝其餘屬性(顏色、線型等)可在這裡補上

                    except AttributeError as attr_err:
                        print(f"跳過無效實體: {entity.ObjectName}, 錯誤: {attr_err}")
                    except Exception as e:
                        print(f"處理實體 {entity.ObjectName} 時發生錯誤: {e}")

                selection_set.Delete()
            except Exception as e:
                print(f"選取或複製過程中發生錯誤: {e}")



        # 定義選取區域範圍
        min_point_a = APoint(100, 150)
        max_point_a = APoint(189.8563, 110)
        copy_selected_objects(acad_doc_a, acad_target_doc, min_point_a, max_point_a, "SS_Temp_A")
        print("a done!")
        min_point_a2 = APoint(37, 80)
        max_point_a2 = APoint(62, 50)
        copy_selected_objects(acad_doc_a, acad_target_doc, min_point_a2, max_point_a2, "SS_Temp_A2")

        min_point_b = APoint(200, 150)
        max_point_b = APoint(234, 78)
        copy_selected_objects(acad_doc_b, acad_target_doc, min_point_b, max_point_b, "SS_Temp_B")

        min_point_b2 = APoint(37, 150)
        max_point_b2 = APoint(72, 80)
        copy_selected_objects(acad_doc_b, acad_target_doc, min_point_b2, max_point_b2, "SS_Temp_B2")

        min_point_b3 = APoint(37, 195)
        max_point_b3 = APoint(65, 172)
        copy_selected_objects(acad_doc_b, acad_target_doc, min_point_b3, max_point_b3, "SS_Temp_B3")

        min_point_c = APoint(100, 130)
        max_point_c = APoint(185, 90)
        copy_selected_objects(acad_doc_c, acad_target_doc, min_point_c, max_point_c, "SS_Temp_C")

        min_point_d = APoint(58.0798, 157)
        max_point_d = APoint(220.2488, 141)
        copy_selected_objects(acad_doc_d, acad_target_doc, min_point_d, max_point_d, "SS_Temp_D")

        # 設置目標圖面中所有物件為黑色
        for entity in acad_target_doc.ModelSpace:
            entity.Color = 7

        # 保存目标图面
        try:
            acad_target_doc.SaveAs(target_path)
        except Exception as e:
            print(f"保存目标文档时发生错误: {e}")

    except Exception as e:
        print(f"操作过程中发生错误: {e}")

    finally:
        # 確保在操作结束后关闭文档
        if 'acad_doc_a' in locals():
            acad_doc_a.Close(False)
        if 'acad_doc_b' in locals():
            acad_doc_b.Close(False)
        if 'acad_doc_c' in locals():
            acad_doc_c.Close(False)
        if 'acad_doc_d' in locals():
            acad_doc_d.Close(False)
        if 'acad_target_doc' in locals():
            acad_target_doc.Close(False)

if __name__ == "__main__":
    merge_drawings_with_specific_areas()



