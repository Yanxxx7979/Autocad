from pyautocad import APoint,Autocad
import math
import pythoncom
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
        if scale_center and scale_factor and scale_factor is not None and scale_center is not None:
            scale_object(ent, APoint(*scale_center), scale_factor)
            # 移動
        if move_delta and move_delta is not None:
            move_object(ent, APoint(*move_delta))


def calculate_x_coordinate(x):
    cad_x=(x-80)*270.27/1595 #80 ->標記圖檔左右留白的寬、270.27/1575 -> cad圖框左右橫向長度/標記圖檔(黑線)左右橫向寬度 、10 ->cad黑框起點
    return cad_x
def calculate_y_coordinate(y):
    cad_x=(1115-(y-60))*189.22/1115 #60 ->標記圖檔上下留白的寬、189.22/1115 -> cad圖框縱向框寬度/標記圖檔(黑線)縱向寬度 、10 ->cad黑框起點 
    return cad_x

def check_center_cood(position, xmin, ymin, xmax, ymax):
    center_p = (xmin+(xmax - xmin) / 2, ymin+(ymin - ymax) / 2, 0)
    print(center_p)
    # 預設 transform_params，避免變數未定義
    transform_params = {
        "move_delta": None,
        "rotate_center": center_p,
        "rotate_angle": None,
        "scale_center": center_p,
        "scale_factor": None
    }
    
    if position == 4:  # 檢查是否等於 '4'
        target_p = (235, 95, 0)
        
        # 正確計算 move_delta（逐項相減）
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 0.8

    elif position == 3:
        target_p = (290,115,0)

        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["rotate_angle"]= math.radians(270)
        transform_params["scale_factor"] = 0.8

    elif position == 2:
        target_p = (35,103,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 0.8

    elif position == 1:
        target_p = (35,140,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 0.8
    
    elif position == 5:
        target_p = (138,180,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 1

    elif position == 6:
        target_p = (138,130,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 1

    elif position == 7:
        target_p = (35,180,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 0.8

    elif position == 8:
        target_p = (303,150,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["rotate_angle"]= math.radians(270)

        transform_params["scale_factor"] = 0.8

    elif position == 9:
        target_p = (118,110,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["move_delta"] = move_delta
        transform_params["scale_factor"] = 0.4

    elif position == 10:
        target_p = (188,95,0)
        
        move_delta = (target_p[0] - center_p[0], target_p[1] - center_p[1], target_p[2] - center_p[2])
        
        print("move_delta:", move_delta)  # 除錯輸出
        
        # 更新 move_delta
        transform_params["rotate_angle"]= math.radians(270)
        transform_params["scale_factor"] = 0.4
        transform_params["move_delta"] = move_delta
    

    return transform_params

    
def process_dwg_within_area(source_file_path, target_file_path, xmin, ymin, xmax, ymax,position):
    """
    將指定範圍內的物件從 source_file_path 複製到 target_file_path。
    範圍由 (xmin, ymin, xmax, ymax) 定義。
    """
    print(xmin,ymin,xmax,ymax)
    acad = Autocad(create_if_not_exists=True)
    transform_params=check_center_cood(position,xmin,ymin,xmax,ymax)
    print(transform_params)
    # 開啟圖面
    try:
        source_doc = acad.Application.Documents.Open(source_file_path)
        target_doc = acad.Application.Documents.Open(target_file_path)
    except Exception as e:
        print(f"無法打開文件: {e}")
        return
    
    # 範圍定義
    min_point = APoint(xmin, ymin)
    max_point = APoint(xmax, ymax)

    def copy_objects_within_area(selection_set_name = "Temp"):
        """複製範圍內的物件到目標圖面"""
        new_created_entities = []  # 儲存新建立的物件
        
      
        for entity in source_doc.ModelSpace:
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
                                
                        elif obj_name == 'AcDbHatch':
                            # Hatch 可能需特別處理
                            # 先嘗試判斷 bounding box
                            try:
                                min_pt_h, max_pt_h = entity.GetBoundingBox()
                                if (min_point.x <= min_pt_h[0] <= max_point.x and
                                    min_point.x <= max_pt_h[0] <= max_point.x and
                                    max_point.y <= min_pt_h[1] <= min_point.y and
                                    max_point.y <= max_pt_h[1] <= min_point.y):
                                    in_range = True
                            except Exception as e:
                                print(f"[Warn] Hatch 無法取得外接矩形: {e}")
                                # 若 bounding box 無法判斷，就暫時跳過
                                pass
                # ------------------------- 通過範圍檢查後 -------------------------                
                # 檢查範圍是否在目標圖面的範圍內
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
                                new_text.ScaleFactor = 0.85

                                if "Arial" in [style.Name for style in target_doc.TextStyles]:
                                    new_text.StyleName = "Arial"
                                else:
                                    style = target_doc.TextStyles.Add("Arial")
                                    style.FontFile = "Arial.ttf"
                                    style.ScaleFactor = 0.85      # 设置宽度因子

                                new_text.StyleName = "Arial"
                                # print("adding")
                                new_created_entities.append(new_text)
                            elif obj_name == 'AcDbMText':
                                new_text = target_doc.ModelSpace.AddMText(
                                    ipt,
                                    entity.height,
                                    entity.TextString
                                )
                                new_text.SetFont("Arial", True, False, 1, 0 or 0)

                                new_text.Layer = layer_name
                                new_text.StyleName = "Arial-Bold"

                                new_text.ScaleFactor = 0.85

                                new_created_entities.append(new_text)
                                # 簡化示範 MText，可自行擴充
                                pass
                            elif obj_name == 'AcDbHatch':
                                # 嘗試直接 CopyObjects
                                hatch_handle = getattr(entity, "Handle", "N/A")
                                print(f"[Info] 找到 AcDbHatch (Handle={hatch_handle}), 嘗試 CopyObjects...")
                                try:
                                    new_objs = source_doc.CopyObjects([entity], target_doc.ModelSpace)
                                    print(f"[Success] 複製 Hatch {hatch_handle} 成功, new_objs數量={len(new_objs)}")
                                    new_created_entities.extend(new_objs)
                                except Exception as copy_err:
                                    print(f"[Error] 複製 Hatch 失敗: {copy_err}")

                    except AttributeError as attr_err:
                        print(f"跳過無效實體: {entity.ObjectName}, 錯誤: {attr_err}")
                    except Exception as e:
                        print(f"處理實體 {entity.ObjectName} 時發生錯誤: {e}")


                selection_set.Delete()

                transform_objects(
                    entities=new_created_entities,
                    move_delta=transform_params["move_delta"],
                    rotate_center=transform_params["rotate_center"],
                    rotate_angle=transform_params["rotate_angle"],
                    scale_center=transform_params["scale_center"],
                    scale_factor=transform_params["scale_factor"]
                )

                return new_created_entities
            except Exception as e:
                print(f"處理選擇集時發生錯誤: {e}")
    # 在transform_objects函式後加上以下代碼來插入文字：
    def add_text_to_target_doc(target_doc, position,height, text_string):
        """
        在目標文件中指定位置添加文字
        """
        # 這裡使用您指定的position座標來設定插入點
        insertion_point = APoint(*position)  # 指定的插入點座標
        
        # 設定文字樣式（可以根據需要調整）
        text_height = height  # 文字的高度
        
        # 在目標文檔的ModelSpace中添加文字
        new_text = target_doc.ModelSpace.AddText(text_string, insertion_point, text_height)
        new_text.Layer = "0"  # 設定文字的圖層（如果需要，可以更改）
        if "標楷體" in [style.Name for style in target_doc.TextStyles]:
            new_text.StyleName = "標楷體"
        else:
            style = target_doc.TextStyles.Add("標楷體")
            style.FontFile = "標楷體.ttf"
            style.ScaleFactor = 0.85   
        # 設定字型，這裡使用Arial字型
   
        new_text.StyleName = "標楷體"

    # 假設您希望在某個特定位置添加文字
    add_text_to_target_doc(target_doc, (105,65,0),5, "Note:")
    add_text_to_target_doc(target_doc, (110,60,0),3, "1.PLUG:")
    copy_objects_within_area(selection_set_name="temp")    

            # 儲存並關閉文件
    if 'source_doc' in locals():
        source_doc.Close(False)
    if 'target_doc' in locals():
        target_doc.Save()  # 或 SaveAs(...)
        target_doc.Close(True)

def annotate_dimension_and_tolerance(xmin, xmax, y_base, tolerance="+0.1/-0.1"):
    pythoncom.CoInitialize()
    acad = Autocad(create_if_not_exists=True)
    model = acad.model

    # 假設標註底部位於圖塊下方 20 單位
    dim_y = y_base - 20

    p1 = APoint(xmin, y_base)
    p2 = APoint(xmax, y_base)
    dim_line_loc = APoint((xmin + xmax) / 2, dim_y)

    # 插入尺寸標註
    dim = model.AddDimAligned(p1, p2, dim_line_loc)

    # 加上公差文字（放在尺寸標註下方一點點）
    tol_text_point = APoint((xmin + xmax) / 2, dim_y - 10)
    model.AddText(f"TOL: {tolerance}", tol_text_point, 5)  # 文字大小你可自訂