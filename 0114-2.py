from pyautocad import Autocad, APoint
import math


def merge_drawings_with_specific_areas():
    # 初始化 AutoCAD
    acad = Autocad(create_if_not_exists=True)


    # 打開圖面文件
    path_a = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\LT-202-UL.dwg"
    path_b = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\LT-501-UL.dwg"
    path_c = r"C:\Users\user\Downloads\Lian Dung\c.dwg"
    path_d = r"C:\Users\user\Downloads\Lian Dung\d.dwg"
    target_path = r"C:\Users\user\Downloads\Lian Dung\Lian Dung\成品圖框(中文版).dwg"


    try:
        acad_doc_a = acad.Application.Documents.Open(path_a)
        acad_doc_b = acad.Application.Documents.Open(path_b)
        acad_doc_c = acad.Application.Documents.Open(path_c)
        acad_doc_d = acad.Application.Documents.Open(path_d)
        acad_target_doc = acad.Application.Documents.Open(target_path)


        # ---------------------------------------------
        # 小工具函式：移動、旋轉、縮放
        # ---------------------------------------------
        def move_object(obj, move_delta):
            if move_delta:
                base_pt = APoint(0, 0, 0)
                obj.Move(base_pt, move_delta)


        def rotate_object(obj, rotate_center, rotate_angle):
            if rotate_center and rotate_angle is not None:
                obj.Rotate(rotate_center, rotate_angle)


        def scale_object(obj, scale_center, scale_factor):
            if scale_center and scale_factor:
                obj.ScaleEntity(scale_center, scale_factor)


        def transform_objects(entities, move_delta=None, rotate_center=None, rotate_angle=None,
                              scale_center=None, scale_factor=None):
            for ent in entities:
                # 旋轉
                if rotate_center and (rotate_angle is not None):
                    rotate_object(ent, APoint(*rotate_center), rotate_angle)
                # 縮放
                if scale_center and scale_factor:
                    scale_object(ent, APoint(*scale_center), scale_factor)
                    # 移動
                if move_delta:
                    move_object(ent, APoint(*move_delta))

        # ---------------------------------------------
        # 核心：篩選 + 複製 (回傳新複製物件清單) + 對新物件做變換
        # ---------------------------------------------
        def copy_selected_objects(source_doc, target_doc,
                                  min_point, max_point, selection_set_name,
                                  move_delta=None, rotate_center=None, rotate_angle=None,
                                  scale_center=None, scale_factor=None):
            """
            1. 在 source_doc.ModelSpace 根據 min_point, max_point 篩選物件 (手動判斷座標)
            2. 建立對應新物件到 target_doc.ModelSpace
            3. 回傳新物件清單
            4. 對新物件做 Move/Rotate/Scale (如有需要)
            """
            new_created_entities = []  # 儲存新建立的物件


            try:
                # 如果選擇集已存在，先刪除
                for selection in source_doc.SelectionSets:
                    if selection.Name == selection_set_name:
                        selection.Delete()
                selection_set = source_doc.SelectionSets.Add(selection_set_name)


                min_x, max_x = min_point.x, max_point.x
                min_y, max_y = min_point.y, max_point.y

                for entity in source_doc.ModelSpace:
                    try:
                        obj_name = entity.ObjectName
                        layer_name = entity.Layer
                        # print(layer_name)

                        # === 1) 檢查圖層是否關閉(Frozen/Off) ===
                        layer_obj = source_doc.Layers.Item(layer_name)
                        # 如果圖層已凍結或關閉，就跳過 (視需求而定)
                        if (not layer_obj.LayerOn) or layer_obj.Freeze:
                            continue


                        # === 2) 檢查圖層名稱，若含有 "center" 則跳過 ===
                        if "center" in layer_name.lower():
                            continue


                        # === 3) 檢查線型是否為 "CENTER" (或包含 "CENTER") ===
                        #       (只有部分物件類型支援 .Linetype，若報錯可改用 entity.LinetypeName)
                        try:
                            linetype_name = entity.Linetype
                            if "center" in linetype_name.lower():
                                continue
                        except:
                            pass  # 有些物件可能取不到線型，忽略即可


                        # === 4) 開始做範圍判斷 ===
                        if obj_name in ['AcDbText', 'AcDbMText', 'AcDbBlockReference']:
                            insertion_point = APoint(*entity.InsertionPoint)
                            # 這裡示範 (min_x <= x <= max_x) & (max_y <= y <= min_y)
                            if not (min_x <= insertion_point.x <= max_x and
                                    max_y <= insertion_point.y <= min_y):
                                continue


                        elif obj_name == 'AcDbLine':
                            start_point = APoint(*entity.StartPoint)
                            end_point   = APoint(*entity.EndPoint)
                            cond1 = (min_x <= start_point.x <= max_x and max_y <= start_point.y <= min_y)
                            cond2 = (min_x <= end_point.x   <= max_x and max_y <= end_point.y   <= min_y)
                            if not (cond1 or cond2):
                                continue


                        elif obj_name == 'AcDbCircle':
                            center_point = APoint(*entity.Center)
                            if not (min_x <= center_point.x <= max_x and
                                    max_y <= center_point.y <= min_y):
                                continue
                        
                        elif obj_name == 'AcDbArc':
                            start_point = APoint(*entity.StartPoint)  # 获取起点
                            end_point = APoint(*entity.EndPoint)      # 获取终点

                            # 检查中心点、起点或终点是否在矩形范围内
                            if not (min_x <= start_point.x <= max_x and
                                    max_y <= start_point.y <= min_y) and \
                            not (min_x <= end_point.x <= max_x and
                                    max_y <= end_point.y <= min_y):
                                # 如果中心点、起点和终点都不在矩形范围内，跳过
                                continue
                                
                        elif entity.ObjectName == "AcDbHatch":
                            # hatch_count += 1
                            hatch_handle = getattr(entity, "Handle", "N/A")  # 取得物件 Handle 以供除錯
                            print(f"[Info] 找到 AcDbHatch (Handle={hatch_handle}). 嘗試 CopyObjects...")

                            # 4) 呼叫 CopyObjects([hatch_entity], target_doc.ModelSpace)
                            try:
                                new_objs = source_doc.CopyObjects([entity], target_doc.ModelSpace)
                                # 若複製成功，可在這裡對 new_objs 做後續處理 (改圖層、顏色…)
                                print(f"[Success] 複製 Hatch {hatch_handle} 成功, new_objs數量={len(new_objs)}")
                            except Exception as copy_err:
                                print(f"[Error] 複製 Hatch {hatch_handle} 失敗: {copy_err}")
                # ------------------------- 通過範圍檢查後 -------------------------                
                # 檢查範圍是否在目標圖面的範圍內
                            except Exception as e:
                                print(f"Failed to calculate bounding box manually: {e}")
                                continue

                        else:
                        # 其他類型暫時不複製
                            continue
                       # ------------------------- 通過範圍檢查後 -------------------------
                        # 確保目標圖面有對應的圖層
                        if layer_name not in target_doc.Layers:
                            target_doc.Layers.Add(layer_name)


                        # 根據類型在目標圖面建立新物件
                        if obj_name == 'AcDbLine':
                            sp = APoint(*entity.StartPoint)
                            ep = APoint(*entity.EndPoint)
                            new_line = target_doc.ModelSpace.AddLine(sp, ep)
                            new_line.Layer = layer_name
                            new_created_entities.append(new_line)


                        elif obj_name == 'AcDbCircle':
                            cpt = APoint(*entity.Center)
                            radius = entity.Radius
                            new_circle = target_doc.ModelSpace.AddCircle(cpt, radius)
                            new_circle.Layer = layer_name
                            new_created_entities.append(new_circle)


                        elif obj_name == 'AcDbArc':
                            cpt = APoint(*entity.Center)
                            radius = entity.Radius
                            start_angle = entity.StartAngle
                            end_angle   = entity.EndAngle
                            new_arc = target_doc.ModelSpace.AddArc(cpt, radius, start_angle, end_angle)
                            new_arc.Layer = layer_name
                            new_created_entities.append(new_arc)


                        elif obj_name in ['AcDbText', 'AcDbMText']:
                            ipt = APoint(*entity.InsertionPoint)
                            if obj_name == 'AcDbText':
                                new_text = target_doc.ModelSpace.AddText(
                                    entity.TextString,
                                    ipt,
                                    entity.Height,  
                      
                                )
                                new_text.Layer = layer_name

                                if "Arial" in [style.Name for style in target_doc.TextStyles]:
                                    new_text.StyleName = "Arial"
                                else:
                                    style = target_doc.TextStyles.Add("Arial")
                                    style.FontFile = "Arial.ttf"
                                    style.WidthFactor = 0.85      # 设置宽度因子

                                new_text.StyleName = "Arial"
                                # print("adding")
                                new_created_entities.append(new_text)
                            elif obj_name == 'AcDbMText':
                                new_text = target_doc.ModelSpace.AddMText(
                                    ipt,
                                    entity.Width,
                                    entity.TextString
                                )
                                new_text.Layer = layer_name
                                new_text.StyleName = "Arial"
                                new_created_entities.append(new_text)
                                # 簡化示範 MText，可自行擴充
                                pass
                            elif entity.ObjectName == "AcDbHatch":
                                print('hatching')
                                try:
                                    # 獲取 Hatch 的相關屬性
                                    pattern_name = entity.PatternName
                                    print(f"Hatch pattern: {pattern_name}")

                                    # 嘗試取得外接矩形
                                    try:
                                        min_pt_h, max_pt_h = entity.GetBoundingBox()
                                        if not (min_x <= min_pt_h[0] <= max_x and 
                                                min_x <= max_pt_h[0] <= max_x and
                                                max_y <= min_pt_h[1] <= min_y and
                                                max_y <= max_pt_h[1] <= min_y):
                                            continue  # 如果不在範圍內則跳過
                                    except Exception as e:
                                        print(f"Failed to get bounding box for hatch: {e}")
                                        continue  # 跳過此 Hatch

                                    # 新增 Hatch 到目標圖面
                                    def add_hatch_to_target(target_doc, layer_name, pattern_name, entity):
                                        # 在目標圖面中新增一個 Hatch
                                        hatch = target_doc.ModelSpace.AddHatch(
                                            PatternType=0,  # 0 表示 Predefined Pattern
                                            PatternName=pattern_name,
                                            Associativity=True
                                        )
                                        # 設定 Hatch 的圖層
                                        hatch.Layer = layer_name

                                        # 嘗試將 Hatch 與邊界關聯 (如果需要)
                                        try:
                                            boundary_entities = entity.GetLoopAt(0)  # 獲取第一個邊界 Loop
                                            for boundary in boundary_entities:
                                                hatch.AppendLoop(0, [boundary])
                                        except Exception as e:
                                            print(f"Failed to associate hatch boundary: {e}")
                                        
                                        return hatch

                                    # 呼叫新增 Hatch 的函數
                                    new_hatch = add_hatch_to_target(target_doc, layer_name, pattern_name, entity)
                                    if new_hatch:
                                        new_created_entities.append(new_hatch)

                                except Exception as e:
                                    print(f"Error processing hatch entity: {e}")

                    except AttributeError as attr_err:
                        print(f"跳過無效實體: {entity.ObjectName}, 錯誤: {attr_err}")
                    except Exception as e:
                        print(f"處理實體 {entity.ObjectName} 時發生錯誤: {e}")


                selection_set.Delete()


                # ★ 對新複製物件進行 Move / Rotate / Scale
                transform_objects(
                    entities       = new_created_entities,
                    move_delta     = move_delta,
                    rotate_center  = rotate_center,
                    rotate_angle   = rotate_angle,
                    scale_center   = scale_center,
                    scale_factor   = scale_factor
                )


                return new_created_entities


            except Exception as e:
                print(f"選取或複製過程中發生錯誤: {e}")
                return []


        # ----------------------------------------------------------
        # 以下範例針對每個範圍做示範：移動、旋轉、縮放
        # ----------------------------------------------------------
        # A 區域：
        min_point_a = APoint(100, 150)
        max_point_a = APoint(195, 110)
        copy_selected_objects(
            source_doc     = acad_doc_a,
            target_doc     = acad_target_doc,
            min_point      = min_point_a,
            max_point      = max_point_a,
            selection_set_name = "SS_Temp_A",
            move_delta     = (-85.7, 50, 0),  
            scale_center   = (0, 0, 0),  
            scale_factor   = 0.8
        )
        print("A done!")
        copy_selected_objects(
            source_doc     = acad_doc_a,
            target_doc     = acad_target_doc,
            min_point      = min_point_a,
            max_point      = max_point_a,
            selection_set_name = "SS_Temp_A",
            move_delta     = (75, 70, 0),  
            scale_center   = (0, 0, 0),  
            scale_factor   = 0.3
        )


        # # # A2 區域：, 再旋轉 0 度 ，以 (0,0) 當旋轉中心
        min_point_a2 = APoint(37, 80)
        max_point_a2 = APoint(62, 50)
        copy_selected_objects(
            source_doc     = acad_doc_a,
            target_doc     = acad_target_doc,
            min_point      = min_point_a2,
            max_point      = max_point_a2,
            selection_set_name = "SS_Temp_A2",
            move_delta     = (-20, 45, 0),    
            rotate_center  = (0, 0, 0),
            rotate_angle   = math.radians(0),  
        )
        print("A2 done!")


        # # B 區域：移動(0,0)，旋轉 180 度，以 (200,150) 為旋轉中心
        min_point_b = APoint(200, 150)
        max_point_b = APoint(235, 75)
        copy_selected_objects(
            source_doc     = acad_doc_b,
            target_doc     = acad_target_doc,
            min_point      = min_point_b,
            max_point      = max_point_b,
            selection_set_name = "SS_Temp_B",
            move_delta     = (87.3, 48, 0),
            rotate_center  = (200, 150, 0),
            rotate_angle   = math.radians(270),
            scale_center   = (50, 0, 0),
            scale_factor   = 0.8
        )
        print("B done!")


        # # B2 區域：同時執行移動 & 旋轉 & 縮放
        min_point_b2 = APoint(37, 150)
        max_point_b2 = APoint(72, 80)
        copy_selected_objects(
            source_doc     = acad_doc_b,
            target_doc     = acad_target_doc,
            min_point      = min_point_b2,
            max_point      = max_point_b2,
            selection_set_name = "SS_Temp_B2",
            move_delta     = (127, 80, 0),
            rotate_center  = (50, 50, 0),
            rotate_angle   = math.radians(270),
            scale_center   = (50, 50, 0),
            scale_factor   = 0.8
        )
        print("B2 done!")

        copy_selected_objects(
            source_doc     = acad_doc_b,
            target_doc     = acad_target_doc,
            min_point      = min_point_b2,
            max_point      = max_point_b2,
            selection_set_name = "SS_Temp_B2",
            move_delta     = (80, 60, 0),
            rotate_center  = (50, 50, 0),
            rotate_angle   = math.radians(270),
            scale_center   = (50, 50, 0),
            scale_factor   = 0.3
        )
        print("B2 done!")


        # # B3 區域：不做變換
        min_point_b3 = APoint(37, 195)
        max_point_b3 = APoint(65, 172)
        copy_selected_objects(
            source_doc     = acad_doc_b,
            target_doc     = acad_target_doc,
            min_point      = min_point_b3,
            max_point      = max_point_b3,
            selection_set_name = "SS_Temp_B3",
            move_delta     = (187, -88, 0),

        )
        print("B3 done!")


        # # # C 區域：只移動
        min_point_c = APoint(100, 130)
        max_point_c = APoint(185, 90)
        copy_selected_objects(
            source_doc     = acad_doc_c,
            target_doc     = acad_target_doc,
            min_point      = min_point_c,
            max_point      = max_point_c,
            selection_set_name = "SS_Temp_C",
            move_delta     = (-6, -5, 0)
        )
        print("C done!")


        # # # D 區域：移動 (-100, 0)
        min_point_d = APoint(58.0798, 157)
        max_point_d = APoint(220.2488, 141)
        copy_selected_objects(
            source_doc     = acad_doc_d,
            target_doc     = acad_target_doc,
            min_point      = min_point_d,
            max_point      = max_point_d,
            selection_set_name = "SS_Temp_D",
            move_delta     = (0, 0, 0)
        )
        print("D done!")


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
