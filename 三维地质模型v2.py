import arcpy
import os

arcpy.env.overwriteOutput = True

# 设置新的基础路径
base_path = r'K:\GIS\Courses\Graduate\GIS2025\GuoJH\GuoJH_三维地质建模v2'

# 设置工作空间
workspace = os.path.join(base_path, 'GuoJH三维地质建模.gdb')
workspace2 = os.path.join(base_path, '三维地质建模2.gdb')

# 定义地层信息映射
layer_info = {
    1: "填土", 2: "粉质黏土", 3: "粗砂", 4: "粉质黏土", 5: "砾砂", 
    6: "黏土", 7: "碎石", 11: "页岩", 12: "石灰岩", 18: "闪长岩", 
    13: "石灰岩", 14: "白云岩", 15: "石灰岩", 17: "白云岩"
}
layer_sequence = [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]

# 检查并创建GDB
def create_gdb_if_not_exists(gdb_path):
    if not arcpy.Exists(gdb_path):
        gdb_dir = os.path.dirname(gdb_path)
        gdb_name = os.path.basename(gdb_path)
        if not os.path.exists(gdb_dir):
            os.makedirs(gdb_dir)
        arcpy.CreateFileGDB_management(gdb_dir, gdb_name)
        print(f"创建地理数据库: {gdb_path}")
    else:
        print(f"地理数据库已存在: {gdb_path}")

# 检查并创建TIN目录
def create_dir_if_not_exists(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
        print(f"创建目录: {dir_path}")
    else:
        print(f"目录已存在: {dir_path}")

# 创建必要的GDB和目录
create_gdb_if_not_exists(workspace)
create_gdb_if_not_exists(workspace2)
create_dir_if_not_exists(os.path.join(base_path, 'TIN'))

arcpy.env.workspace = workspace

# 创建表格
table_name = '钻孔地层信息'
arcpy.CreateTable_management(workspace, table_name)

# 添加字段
field_names = ['钻孔编号', '地层编号', '地层名称', '地表高程', '底层埋深', 'X坐标', 'Y坐标']
field_types = ['TEXT', 'DOUBLE', 'TEXT', 'DOUBLE', 'DOUBLE', 'DOUBLE', 'DOUBLE']
for field_name, field_type in zip(field_names, field_types):
    # 修正：使用完整路径添加字段
    arcpy.AddField_management(os.path.join(workspace, table_name), field_name, field_type)

# 插入数据
with arcpy.da.InsertCursor(table_name, field_names) as cursor:
    # 钻孔编号列表
    drillholes = ['ZK2', 'ZK6', 'ZK10', 'ZK16', 'ZK20', 'ZK25', 'ZK29', 'ZK33', 'ZK37', 'ZK41', 'ZK48', 
                  'ZK53', 'ZK56', 'ZK59', 'ZK62', 'ZK64', 'ZK65', 'ZK67', 'ZK69', 'ZK71', 'ZK73', 'ZK78', 
                  'ZK80', 'ZK82', 'ZK84', 'ZK86', 'ZK89', 'ZK92', 'ZK95', 'ZK98', 'ZK100', 'ZK103']
    
    # 地层编号列表
    layer_numbers = [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]
    
    # 循环每个钻孔
    for drillhole in drillholes:
        for layer_number in layer_numbers:
            # 填充数据
            cursor.insertRow([drillhole, layer_number, '', None, None, None, None])

print("Table created successfully!")


# 设置输入表格和游标
excel_table = 'ExcelTable'
layer_info_table = '钻孔地层信息'
fields_to_update = ['地层名称', '地表高程', '底层埋深', 'X坐标', 'Y坐标']

# 创建Excel到GDB表
excel_file = os.path.join(base_path, "钻孔总表.xlsx")
excel_to_table = 'ExcelTable'
arcpy.ExcelToTable_conversion(excel_file, os.path.join(workspace, excel_to_table))

# 创建更新游标
with arcpy.da.UpdateCursor(layer_info_table, ['钻孔编号', '地层编号'] + fields_to_update) as update_cursor:
    # 构建钻孔编号和地层编号的索引，以便快速检索
    excel_index = {}
    with arcpy.da.SearchCursor(excel_table, ['钻孔编号', '地层编号'] + fields_to_update) as search_cursor:
        for row in search_cursor:
            zk_bh, dc_bh = row[:2]
            excel_index[(zk_bh, dc_bh)] = row[2:]
    
    # 遍历钻孔地层信息表
    for row in update_cursor:
        zk_bh, dc_bh = row[:2]
        if (zk_bh, dc_bh) in excel_index:
            # 如果在Excel表中找到匹配的记录，则更新相应字段
            row[2:] = excel_index[(zk_bh, dc_bh)]
            update_cursor.updateRow(row)

print("Update completed successfully!")


# 钻孔编号列表
drillhole_list = [
    "ZK2", "ZK6", "ZK10", "ZK16", "ZK20", "ZK25", "ZK29", "ZK33", "ZK37", "ZK41",
    "ZK48", "ZK53", "ZK56", "ZK59", "ZK62", "ZK64", "ZK65", "ZK67", "ZK69", "ZK71",
    "ZK73", "ZK78", "ZK80", "ZK82", "ZK84", "ZK86", "ZK89", "ZK92", "ZK95", "ZK98",
    "ZK100", "ZK103"
]

# 地层顺序列表
all_layers = [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]

# 地层名称映射
layer_name = {
    1: "填土", 2: "粉质黏土", 3: "粗砂", 4: "粉质黏土", 5: "砾砂", 6: "黏土", 7: "碎石",
    11: "页岩", 12: "石灰岩", 18: "闪长岩", 13: "石灰岩", 14: "白云岩", 15: "石灰岩", 17: "白云岩"
}


for drillhole in drillhole_list:
    # 初始化地层编号和相关信息的字典
    real_layers = {}
    virtual_layers = [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]
    print(f'当前钻孔：{drillhole}')
    
    # 使用searchCursor寻找非空地层信息
    with arcpy.da.SearchCursor("钻孔地层信息", ["地层编号", "地层名称", "地表高程", "底层埋深", "X坐标", "Y坐标"], f"钻孔编号 = '{drillhole}'") as sCursor:
        for row in sCursor:            
            if row[1]:  # 地层名称非空
                real_layers[row[0]] = {"地层名称": row[1], "地表高程": row[2], "底层埋深": row[3], "X坐标": row[4], "Y坐标": row[5]}
                virtual_layers.remove(row[0])
    
    # 更新缺失的地层信息
    with arcpy.da.UpdateCursor("钻孔地层信息", ["地层编号", "地层名称", "地表高程", "底层埋深", "X坐标", "Y坐标"], f"钻孔编号 = '{drillhole}'") as uCursor:
        for row in uCursor:
            print(f'...当前地层：{int(row[0])}')
            if row[0] in virtual_layers:
                print(f'......插入虚拟地层：{int(row[0])}')
                row[1] = layer_name[row[0]]
                first_real_layers_key = list(real_layers.keys())[0]
                first_real_layer = real_layers[first_real_layers_key]  # 获取第一个实际地层的信息
                if row[0] == 1:  # 如果地层1就是虚拟地层
                    row[2] = first_real_layer["地表高程"]
                    row[3] = 0.01  # 底层埋深设为0.01
                    row[4] = first_real_layer["X坐标"]
                    row[5] = first_real_layer["Y坐标"]
                    real_layers[1] =  {"地层名称": row[1], "地表高程": row[2], "底层埋深": row[3], "X坐标": row[4], "Y坐标": row[5]}
                else:  # 根据前一地层信息更新
                    prev_layer_index = all_layers.index(row[0]) - 1
                    prev_layer = all_layers[prev_layer_index]
                    row[2] = real_layers[prev_layer]["地表高程"]
                    row[3] = real_layers[prev_layer]["底层埋深"] 
                    row[4] = real_layers[prev_layer]["X坐标"]
                    row[5] = real_layers[prev_layer]["Y坐标"]
                    real_layers[row[0]] =  {"地层名称": row[1], "地表高程": row[2], "底层埋深": row[3], "X坐标": row[4], "Y坐标": row[5]}
                uCursor.updateRow(row)
               


# 定义要操作的表格路径
table_path = os.path.join(workspace, '钻孔地层信息')

# 添加新字段"底层高程"
arcpy.AddField_management(table_path, "底层高程", "DOUBLE")

# 使用UpdateCursor更新新字段的值
with arcpy.da.UpdateCursor(table_path, ["地表高程", "底层埋深", "底层高程"]) as cursor:
    for row in cursor:
        # 计算"底层高程" = "地表高程" - "底层埋深"
        row[2] = row[0] - row[1]
        cursor.updateRow(row)

print("字段添加并更新完成。")
#==================生成真实钻孔坐标点====================================================
print("\n开始生成真实钻孔点...")

try:
    # 设置空间参考
    sr = arcpy.SpatialReference(4548)  # CGCS2000
    
    # 创建点要素类
    output_name = "layer_0"
    
    # 如果已存在则删除
    if arcpy.Exists(output_name):
        arcpy.Delete_management(output_name)
    
    # 创建带Z值的点要素类
    arcpy.CreateFeatureclass_management(
        out_path=arcpy.env.workspace,
        out_name=output_name,
        geometry_type="POINT",
        spatial_reference=sr,
        has_z="ENABLED"
    )
    
    # 添加属性字段
    # 修正：使用完整路径添加字段
    arcpy.AddField_management(os.path.join(workspace, output_name), "钻孔编号", "TEXT", field_length=50)
    arcpy.AddField_management(os.path.join(workspace, output_name), "地层编号", "LONG")
    arcpy.AddField_management(os.path.join(workspace, output_name), "地层名称", "TEXT", field_length=50)
    arcpy.AddField_management(os.path.join(workspace, output_name), "地表高程", "DOUBLE")
    
    # 插入所有钻孔的所有地层点
    print("正在创建钻孔点...")
    with arcpy.da.InsertCursor(output_name, ["SHAPE@XY", "SHAPE@Z", "钻孔编号", "地层编号", "地层名称", "地表高程"]) as icursor:
        # 从钻孔地层信息表读取所有记录
        with arcpy.da.SearchCursor("钻孔地层信息", ["钻孔编号", "地层编号", "地层名称", "地表高程", "X坐标", "Y坐标"]) as scursor:
            for row in scursor:
                hole_id, layer_id, layer_name, elevation, x, y = row
                if None not in (x, y, elevation):  # 确保坐标和高程值不为空
                    point_xy = (x, y)
                    icursor.insertRow([point_xy, elevation, hole_id, layer_id, layer_name, elevation])
    
    # 统计结果
    point_count = int(arcpy.GetCount_management(output_name)[0])
    
    # 获取唯一钻孔数量
    unique_holes = set()
    with arcpy.da.SearchCursor(output_name, ["钻孔编号"]) as cursor:
        for row in cursor:
            unique_holes.add(row[0])
    
    print(f"成功创建钻孔点要素类！")
    print(f"  总计创建点数量: {point_count}")
    print(f"  涉及钻孔数量: {len(unique_holes)}")
    print(f"  要素类保存位置: {arcpy.env.workspace}")

except arcpy.ExecuteError as e:
    print("处理过程中发生ArcGIS错误：")
    print(arcpy.GetMessages(2))
except Exception as e:
    print(f"发生未知错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
#========================================为虚拟钻孔创建泰森多边形==============================================
import arcpy
import os
import numpy as np

# 设置环境和输入数据
arcpy.env.overwriteOutput = True
# 修改输入路径和工作空间路径
input_points = os.path.join(workspace, "layer_0") 

# ...其余代码保持不变...
output_thiessen = os.path.join(workspace, "thiessen_polygons")
output_vertices = os.path.join(workspace, "virtual_boreholes")

# 步骤1：创建泰森多边形
arcpy.analysis.CreateThiessenPolygons(
    in_features=input_points,
    out_feature_class=output_thiessen,
    fields_to_copy="ALL"  # 包含所有原始字段
)
print("泰森多边形创建完成")

# 步骤2：提取泰森多边形的顶点作为点
# 创建一个新的要素类来存储顶点
arcpy.CreateFeatureclass_management(
    out_path=workspace,
    out_name=os.path.basename(output_vertices),
    geometry_type="POINT",
    spatial_reference=arcpy.Describe(output_thiessen).spatialReference
)

# 为虚拟钻孔添加属性字段，用于存储关联的原始钻孔ID
# 修正：使用完整路径添加字段
arcpy.AddField_management(os.path.join(workspace, output_vertices), "Related_ID1", "LONG")
arcpy.AddField_management(os.path.join(workspace, output_vertices), "Related_ID2", "LONG")
arcpy.AddField_management(os.path.join(workspace, output_vertices), "Related_ID3", "LONG")  # 添加第三个关联ID字段
arcpy.AddField_management(os.path.join(workspace, output_vertices), "Num_Related", "SHORT")  # 记录关联的多边形数量
arcpy.AddField_management(os.path.join(workspace, output_vertices), "Vertex_ID", "LONG")

# 用于存储已处理的顶点坐标，避免重复
processed_vertices = set()
vertex_id = 1

# 处理每个泰森多边形，提取顶点
# 注意：现在使用Input_FID而不是OBJECTID来关联原始钻孔
with arcpy.da.SearchCursor(output_thiessen, ["SHAPE@", "Input_FID"]) as search_cur:
    with arcpy.da.InsertCursor(output_vertices, ["SHAPE@", "Related_ID1", "Related_ID2", "Related_ID3", "Num_Related", "Vertex_ID"]) as insert_cur:
        
        # 获取所有多边形及其Input_FID，用于后续处理
        polygons = []
        for row in search_cur:
            polygons.append((row[0], row[1]))  # row[1]现在是Input_FID
        
        # 处理每个多边形的顶点
        for i, (polygon, input_fid) in enumerate(polygons):
            
            # 获取多边形的顶点
            for part in polygon:
                for pnt in part:
                    if pnt:  # 跳过None点（内部环起点）
                        # 使用坐标的字符串表示作为唯一标识符
                        vertex_key = f"{pnt.X:.6f},{pnt.Y:.6f}"
                        
                        if vertex_key not in processed_vertices:
                            processed_vertices.add(vertex_key)
                            
                            # 查找与此顶点相邻的所有多边形，记录它们的Input_FID
                            related_ids = [input_fid]  # 使用Input_FID而不是polygon_id
                            for j, (other_polygon, other_input_fid) in enumerate(polygons):
                                if i != j:  # 不与自身比较
                                    # 使用一个很小的容差检查顶点是否属于其他多边形
                                    if other_polygon.distanceTo(arcpy.Point(pnt.X, pnt.Y)) < 0.001:
                                        related_ids.append(other_input_fid)
                                        # 不要break，继续查找所有相邻多边形
                            
                            # 记录关联的多边形数量
                            num_related = len(related_ids)
                            
                            # 确保有三个关联ID槽位
                            related_id1 = related_ids[0] if len(related_ids) > 0 else None
                            related_id2 = related_ids[1] if len(related_ids) > 1 else None
                            related_id3 = related_ids[2] if len(related_ids) > 2 else None
                            
                            # 插入顶点作为点
                            point = arcpy.Point(pnt.X, pnt.Y)
                            insert_cur.insertRow([point, related_id1, related_id2, related_id3, num_related, vertex_id])
                            vertex_id += 1

print("虚拟钻孔（泰森多边形顶点）已创建")

# 统计结果信息
with arcpy.da.SearchCursor(output_vertices, ["Num_Related"]) as cursor:
    counts = {}
    for row in cursor:
        num = row[0]
        if num in counts:
            counts[num] += 1
        else:
            counts[num] = 1
    
    print("虚拟钻孔关联的实际钻孔数量统计:")
    for num, count in sorted(counts.items()):
        print(f"  关联 {num} 个实际钻孔的虚拟钻孔: {count} 个")

print("处理完成！")
#========================================为虚拟钻孔创建表格==============================================
# 创建虚拟钻孔地层表
virtual_strata_table = "virtual_borehole_strata"
arcpy.CreateTable_management(workspace, virtual_strata_table)

# 添加字段
# 修正：使用完整路径添加字段
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "钻孔编号", "TEXT", field_length=50)
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "地层编号", "LONG")
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "地层名称", "TEXT", field_length=50)
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "地表高程", "DOUBLE")
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "底层埋深", "DOUBLE")
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "X坐标", "DOUBLE")
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "Y坐标", "DOUBLE")
arcpy.AddField_management(os.path.join(workspace, virtual_strata_table), "地层厚度", "DOUBLE")

# 获取所有虚拟钻孔的基本信息
virtual_holes = {}
with arcpy.da.SearchCursor(output_vertices, ["Vertex_ID", "SHAPE@XY"]) as cursor:
    for row in cursor:
        vertex_id = row[0]
        x, y = row[1]
        virtual_holes[vertex_id] = {"X": x, "Y": y}

# 插入数据
with arcpy.da.InsertCursor(virtual_strata_table, 
                          ["钻孔编号", "地层编号", "地层名称", "地表高程", 
                           "底层埋深", "X坐标", "Y坐标", "地层厚度"]) as cursor:
    
    # 为每个虚拟钻孔生成地层数据
    for vertex_id in virtual_holes:
        hole_name = f"V{vertex_id}"  # 虚拟钻孔编号格式：V1, V2, ...
        x = virtual_holes[vertex_id]["X"]
        y = virtual_holes[vertex_id]["Y"]
        
        # 初始化当前钻孔的累计深度
        prev_depth = 0
        
        # 为每个地层生成记录
        for layer_num in layer_sequence:
            # 模拟地层深度（这里使用简单的累加方式，可以根据需要修改）
            if layer_num == 1:
                depth = 2.0  # 第一层深度
            else:
                depth = prev_depth + 1.5  # 每层增加1.5米深度
            
            # 计算地层厚度
            thickness = depth - prev_depth
            
            # 插入记录
            cursor.insertRow([
                hole_name,                # 钻孔编号
                layer_num,                # 地层编号
                layer_info[layer_num],    # 地层名称
                100.0,                    # 地表高程（示例值，需要根据实际情况修改）
                depth,                    # 底层埋深
                x,                        # X坐标
                y,                        # Y坐标
                thickness                 # 地层厚度
            ])
            
            # 更新前一层深度
            prev_depth = depth

print("虚拟钻孔地层数据表创建完成！")
#==================================为实际钻孔数据表添加地层厚度字段============================
# ...existing code...

# 添加地层厚度字段
arcpy.AddField_management(table_path, "地层厚度", "DOUBLE")

# 定义地层顺序
layer_sequence = [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]

# 使用字典存储每个钻孔的地层信息
drill_layers = {}

# 首先收集所有钻孔的地层信息
with arcpy.da.SearchCursor(table_path, ["钻孔编号", "地层编号", "底层埋深"]) as cursor:
    for row in cursor:
        drill_hole, layer_num, depth = row
        if drill_hole not in drill_layers:
            drill_layers[drill_hole] = {}
        drill_layers[drill_hole][layer_num] = depth

# 更新地层厚度
with arcpy.da.UpdateCursor(table_path, ["钻孔编号", "地层编号", "底层埋深", "地层厚度"]) as cursor:
    for row in cursor:
        drill_hole, layer_num, depth, _ = row
        
        if layer_num == 1:  # 对于第一层
            thickness = depth
        else:
            # 找到当前地层在序列中的位置
            current_index = layer_sequence.index(layer_num)
            # 获取上一层的编号
            prev_layer = layer_sequence[current_index - 1]
            
            # 获取上一层的底层埋深
            prev_depth = drill_layers[drill_hole].get(prev_layer, 0)
            thickness = depth - prev_depth
        
        row[3] = thickness  # 更新地层厚度
        cursor.updateRow(row)

print("地层厚度计算完成。")
#=======================为虚拟钻孔更新数据（地层厚度和地层埋深）=========================================

print("\n开始更新虚拟钻孔数据...")

try:
    # 获取实际钻孔数据
    print("读取实际钻孔数据...")
    real_holes_data = {}
    with arcpy.da.SearchCursor("钻孔地层信息", 
                              ["钻孔编号", "地层编号", "底层埋深", "地层厚度"]) as cursor:
        for row in cursor:
            hole_id = row[0]
            if hole_id not in real_holes_data:
                real_holes_data[hole_id] = {}
            real_holes_data[hole_id][row[1]] = {
                "depth": row[2],
                "thickness": row[3]
            }
    
    # 获取虚拟钻孔关联关系
    print("读取虚拟钻孔关联关系...")
    virtual_relations = {}
    with arcpy.da.SearchCursor(output_vertices, 
                              ["Vertex_ID", "Related_ID1", "Related_ID2", "Related_ID3"]) as cursor:
        for row in cursor:
            vertex_id = f"V{row[0]}"
            # 获取关联的实际钻孔编号
            related_holes = []
            for related_id in row[1:]:
                if related_id is not None:
                    with arcpy.da.SearchCursor("layer_0", 
                                             ["钻孔编号"], 
                                             f"OBJECTID = {related_id}") as hole_cursor:
                        for hole_row in hole_cursor:
                            related_holes.append(hole_row[0])
            virtual_relations[vertex_id] = related_holes
    
    # 更新虚拟钻孔数据
    total_holes = len(virtual_relations)
    print(f"\n开始更新 {total_holes} 个虚拟钻孔的数据...")
    
    update_count = 0
    zero_thickness_count = 0
    current_hole = 0
    
    for virtual_hole, related_holes in virtual_relations.items():
        current_hole += 1
        print(f"\r正在处理: {virtual_hole} ({current_hole}/{total_holes})", end="")
        
        prev_depth = 0
        for layer_num in layer_sequence:
            # 收集关联钻孔的地层数据
            depths = []
            has_zero_thickness = False
            
            for real_hole in related_holes:
                if real_hole in real_holes_data and layer_num in real_holes_data[real_hole]:
                    hole_data = real_holes_data[real_hole][layer_num]
                    if hole_data["thickness"] == 0:
                        has_zero_thickness = True
                        break
                    depths.append(hole_data["depth"])
            
            # 更新当前地层
            where_clause = f"钻孔编号 = '{virtual_hole}' AND 地层编号 = {layer_num}"
            with arcpy.da.UpdateCursor(virtual_strata_table, 
                                     ["底层埋深", "地层厚度"], 
                                     where_clause) as update_cursor:
                for row in update_cursor:
                    if has_zero_thickness:
                        if layer_num == 1:
                            row[0] = 0
                        else:
                            row[0] = prev_depth
                        row[1] = 0
                        zero_thickness_count += 1
                    elif depths:
                        avg_depth = sum(depths) / len(depths)
                        row[0] = avg_depth
                        row[1] = avg_depth - prev_depth if layer_num > 1 else avg_depth
                    
                    prev_depth = row[0]
                    update_cursor.updateRow(row)
                    update_count += 1
    
    print(f"\n\n更新完成！")
    print(f"更新记录数: {update_count}")
    print(f"零厚度地层数: {zero_thickness_count}")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())

#=======================为虚拟钻孔更新数据（地表高程）===============================================
print("\n开始更新虚拟钻孔地表高程...")

try:
    # 获取实际钻孔地表高程
    print("读取实际钻孔地表高程...")
    real_holes_elevation = {}
    with arcpy.da.SearchCursor("钻孔地层信息", ["钻孔编号", "地表高程"]) as cursor:
        for row in cursor:
            if row[0] not in real_holes_elevation:
                real_holes_elevation[row[0]] = row[1]
    
    # 获取虚拟钻孔关联关系
    print("读取虚拟钻孔关联关系...")
    virtual_relations = {}
    with arcpy.da.SearchCursor(output_vertices, 
                              ["Vertex_ID", "Related_ID1", "Related_ID2", "Related_ID3"]) as cursor:
        for row in cursor:
            vertex_id = f"V{row[0]}"
            related_holes = []
            for related_id in row[1:]:
                if related_id is not None:
                    with arcpy.da.SearchCursor("layer_0", 
                                             ["钻孔编号"], 
                                             f"OBJECTID = {related_id}") as hole_cursor:
                        for hole_row in hole_cursor:
                            related_holes.append(hole_row[0])
            virtual_relations[vertex_id] = related_holes
    
    # 更新虚拟钻孔地表高程
    total_holes = len(virtual_relations)
    print(f"\n开始更新 {total_holes} 个虚拟钻孔的地表高程...")
    update_count = 0
    current_hole = 0
    
    # 使用集合来追踪已处理的钻孔
    processed_holes = set()
    
    for virtual_hole, related_holes in virtual_relations.items():
        current_hole += 1
        print(f"\r正在处理: {virtual_hole} ({current_hole}/{total_holes})", end="")
        
        # 收集关联钻孔的地表高程
        elevations = []
        for real_hole in related_holes:
            if real_hole in real_holes_elevation:
                elevation = real_holes_elevation[real_hole]
                if elevation is not None:
                    elevations.append(elevation)
        
        # 更新虚拟钻孔的所有地层记录的地表高程
        if elevations:
            avg_elevation = sum(elevations) / len(elevations)
            where_clause = f"钻孔编号 = '{virtual_hole}'"
            
            with arcpy.da.UpdateCursor(virtual_strata_table, 
                                     ["钻孔编号", "地表高程"], 
                                     where_clause) as update_cursor:
                for row in update_cursor:
                    if row[0] not in processed_holes:
                        row[1] = avg_elevation
                        update_cursor.updateRow(row)
                        update_count += 1
                        processed_holes.add(row[0])
    
    print(f"\n\n地表高程更新完成！")
    print(f"更新钻孔数量: {len(processed_holes)}")
    print(f"总更新记录数: {update_count}")

    # 验证结果
    with arcpy.da.SearchCursor(virtual_strata_table, ["地表高程"]) as cursor:
        elevations = [row[0] for row in cursor]
        if elevations:
            print(f"\n地表高程统计:")
            print(f"  最小值: {min(elevations):.2f}")
            print(f"  最大值: {max(elevations):.2f}")
            print(f"  平均值: {sum(elevations)/len(elevations):.2f}")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
    #=======================为其他地层更新地表高程================================================
print("\n开始更新虚拟钻孔地表高程...")

try:
    # 获取编号为1的地层的地表高程
    print("读取第一层地表高程...")
    virtual_hole_elevations = {}
    with arcpy.da.SearchCursor(virtual_strata_table, 
                              ["钻孔编号", "地表高程"], 
                              "地层编号 = 1") as cursor:
        for row in cursor:
            virtual_hole_elevations[row[0]] = row[1]
    
    print(f"共获取到 {len(virtual_hole_elevations)} 个虚拟钻孔的地表高程")
    
    # 更新所有地层的地表高程
    update_count = 0
    with arcpy.da.UpdateCursor(virtual_strata_table, 
                              ["钻孔编号", "地表高程"]) as cursor:
        for row in cursor:
            hole_id = row[0]
            if hole_id in virtual_hole_elevations:
                row[1] = virtual_hole_elevations[hole_id]
                cursor.updateRow(row)
                update_count += 1
    
    print("\n地表高程更新完成！")
    print(f"更新记录数: {update_count}")

    # 验证结果
    verification_count = 0
    with arcpy.da.SearchCursor(virtual_strata_table, 
                              ["钻孔编号", "地表高程", "地层编号"]) as cursor:
        prev_hole = None
        prev_elevation = None
        for row in cursor:
            hole_id, elevation, layer_id = row
            if prev_hole != hole_id:
                prev_hole = hole_id
                prev_elevation = elevation
            else:
                # 检查同一钻孔的地表高程是否相同
                if elevation != prev_elevation:
                    print(f"警告: 钻孔 {hole_id} 的地层 {layer_id} 地表高程不一致")
                    verification_count += 1
    
    if verification_count == 0:
        print("验证通过：所有同一钻孔的地层共用相同的地表高程")
    else:
        print(f"发现 {verification_count} 处高程不一致")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
    #=======================合并钻孔数据表===========================================================
print("\n开始合并钻孔数据表...")

try:
    # 定义输出表名
    merged_table = "merged_borehole_strata"
    
    # 创建新表
    print("创建合并后的数据表...")
    arcpy.CreateTable_management(workspace2, merged_table)
    
    # 添加字段
    print("添加字段...")
    # 修正：使用完整路径添加字段
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "钻孔编号", "TEXT", field_length=50)
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "地层编号", "LONG")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "地层名称", "TEXT", field_length=50)
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "地表高程", "DOUBLE")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "底层埋深", "DOUBLE")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "X坐标", "DOUBLE")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "Y坐标", "DOUBLE")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "地层厚度", "DOUBLE")
    arcpy.AddField_management(os.path.join(workspace2, merged_table), "钻孔类型", "TEXT", field_length=10)
    
    # 插入实际钻孔数据
    print("\n导入实际钻孔数据...")
    real_count = 0
    fields = ["钻孔编号", "地层编号", "地层名称", "地表高程", "底层埋深", "X坐标", "Y坐标", "地层厚度"]
    all_fields = fields + ["钻孔类型"]
    
    with arcpy.da.InsertCursor(os.path.join(workspace2, merged_table), all_fields) as insert_cur:
        with arcpy.da.SearchCursor(os.path.join(workspace, "钻孔地层信息"), fields) as search_cur:
            for row in search_cur:
                insert_cur.insertRow(list(row) + ["实际"])
                real_count += 1
    
    # 插入虚拟钻孔数据
    print("导入虚拟钻孔数据...")
    virtual_count = 0
    with arcpy.da.InsertCursor(os.path.join(workspace2, merged_table), all_fields) as insert_cur:
        with arcpy.da.SearchCursor(os.path.join(workspace, virtual_strata_table), fields) as search_cur:
            for row in search_cur:
                insert_cur.insertRow(list(row) + ["虚拟"])
                virtual_count += 1
    
    # 验证结果
    total_count = int(arcpy.GetCount_management(os.path.join(workspace2, merged_table))[0])
    print("\n合并完成！")
    print(f"实际钻孔记录数: {real_count}")
    print(f"虚拟钻孔记录数: {virtual_count}")
    print(f"总记录数: {total_count}")
    
    # 统计唯一钻孔数量
    unique_holes = {"实际": set(), "虚拟": set()}
    with arcpy.da.SearchCursor(os.path.join(workspace2, merged_table), ["钻孔编号", "钻孔类型"]) as cursor:
        for row in cursor:
            unique_holes[row[1]].add(row[0])
    
    print(f"\n钻孔统计:")
    print(f"  实际钻孔数量: {len(unique_holes['实际'])}")
    print(f"  虚拟钻孔数量: {len(unique_holes['虚拟'])}")
    print(f"  总钻孔数量: {len(unique_holes['实际'] | unique_holes['虚拟'])}")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
    #=======================添加并计算底层高程===========================================================
print("\n开始添加并计算底层高程...")

try:
    table_name = merged_table
    
    # 检查字段是否已存在
    existing_fields = [field.name for field in arcpy.ListFields(os.path.join(workspace2, table_name))]
    if "底层高程" not in existing_fields:
        print("添加底层高程字段...")
        arcpy.AddField_management(os.path.join(workspace2, table_name), "底层高程", "DOUBLE")
    
    # 更新底层高程
    print("计算底层高程...")
    update_count = 0
    with arcpy.da.UpdateCursor(os.path.join(workspace2, table_name), 
                              ["地表高程", "底层埋深", "底层高程"]) as cursor:
        for row in cursor:
            surface_elevation = row[0]
            burial_depth = row[1]
            
            # 计算底层高程
            if surface_elevation is not None and burial_depth is not None:
                row[2] = surface_elevation - burial_depth
                cursor.updateRow(row)
                update_count += 1
    
    print(f"底层高程计算完成！")
    print(f"更新记录数: {update_count}")
    
    # 验证结果
    with arcpy.da.SearchCursor(os.path.join(workspace2, table_name), ["底层高程"]) as cursor:
        elevations = [row[0] for row in cursor if row[0] is not None]
        if elevations:
            print("\n底层高程统计:")
            print(f"  最小值: {min(elevations):.2f}")
            print(f"  最大值: {max(elevations):.2f}")
            print(f"  平均值: {sum(elevations)/len(elevations):.2f}")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
    #==========================为最底层地层更新底层埋深，使整个模型更加规整============================
print("\n开始更新最底层地层底层埋深...")

try:
    # 设置工作空间和表名
    table_name = merged_table
    
    # 创建字典存储每个钻孔的最底层有效地层信息
    bottom_layers = {}
    max_depth = 0
    
    # 第一次遍历：找出每个钻孔的最底层有效地层（厚度不为0）
    print("识别最底层有效地层...")
    with arcpy.da.SearchCursor(os.path.join(workspace2, table_name), 
                             ["钻孔编号", "地层编号", "底层埋深", "地层厚度"]) as cursor:
        for row in cursor:
            hole_id, layer_id, depth, thickness = row
            # 只考虑厚度不为0的地层
            if thickness and thickness > 0:
                if hole_id not in bottom_layers:
                    bottom_layers[hole_id] = {"layer": layer_id, "depth": depth}
                elif layer_id > bottom_layers[hole_id]["layer"]:
                    bottom_layers[hole_id] = {"layer": layer_id, "depth": depth}
                
                # 同时记录所有有效地层中最大的底层埋深
                if depth and depth > max_depth:
                    max_depth = depth
    
    print(f"找到 {len(bottom_layers)} 个钻孔的最底层有效地层")
    print(f"最大底层埋深: {max_depth:.2f}")
    
    # 第二次遍历：更新最底层有效地层的底层埋深
    print("\n开始更新底层埋深...")
    update_count = 0
    with arcpy.da.UpdateCursor(os.path.join(workspace2, table_name), 
                             ["钻孔编号", "地层编号", "底层埋深", "地层厚度"]) as cursor:
        for row in cursor:
            hole_id, layer_id, depth, thickness = row
            # 检查是否是该钻孔的最底层有效地层
            if (hole_id in bottom_layers and 
                layer_id == bottom_layers[hole_id]["layer"] and 
                thickness and thickness > 0):
                # 更新为最大底层埋深
                row[2] = max_depth
                cursor.updateRow(row)
                update_count += 1
    
    print(f"更新完成！共更新 {update_count} 条记录")
    
    # 验证结果
    print("\n验证结果...")
    depths = []
    with arcpy.da.SearchCursor(os.path.join(workspace2, table_name), ["底层埋深", "地层厚度"]) as cursor:
        for row in cursor:
            if row[1] and row[1] > 0:  # 只统计有效地层
                depths.append(row[0])
    
    if depths:
        print(f"底层埋深统计:")
        print(f"  最小值: {min(depths):.2f}")
        print(f"  最大值: {max(depths):.2f}")
        print(f"  平均值: {sum(depths)/len(depths):.2f}")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
    
#=========================================创建tin==========================================
# 设置工作环境
arcpy.env.overwriteOutput = True
arcpy.env.workspace = workspace2

# 创建TIN目录
tin_dir = os.path.join(base_path, 'TIN')
if not os.path.exists(tin_dir):
    os.makedirs(tin_dir)

# 使用指定的坐标系统WKID 4548
sr = arcpy.SpatialReference(4548)

# 地层编号列表，包括地表层0
layers = [0] + [1, 2, 3, 4, 5, 6, 7, 11, 12, 18, 13, 14, 15, 17]

# 准备字段名
fields = ["X坐标", "Y坐标", "地表高程", "底层埋深"]

# 生成点要素类并创建TIN
for layer_id in layers:
    # 定义输出点要素类和TIN的路径
    output_fc_name = f"layer_{layer_id}"
    output_fc_path = os.path.join(workspace2, output_fc_name)
    tin_name = f"TIN_{layer_id}"
    tin_path = os.path.join(tin_dir, tin_name)

    # 创建点要素类
    arcpy.CreateFeatureclass_management(out_path=workspace2, out_name=output_fc_name, geometry_type="POINT", spatial_reference=sr, has_z="ENABLED")

    # 插入点
    with arcpy.da.InsertCursor(output_fc_path, ["SHAPE@XY", "SHAPE@Z"]) as icursor:
        query = "地层编号 = 1" if layer_id == 0 else f"地层编号 = {layer_id}"
        z_field = "地表高程" if layer_id == 0 else "底层埋深"
        for row in arcpy.da.SearchCursor(os.path.join(workspace2, merged_table), fields, query):
            point = arcpy.Point(row[0], row[1])
            z_value = row[2] if layer_id == 0 else row[2] - row[3]  # 使用地表高程减去底层埋深获取Z值
            icursor.insertRow([(point.X, point.Y), z_value])

    # 生成TIN
    arcpy.CreateTin_3d(tin_path, sr, f'{output_fc_path} Shape.Z Mass_Points <None>', "DELAUNAY")

    print(f"{tin_name} 创建完成。")
#=======================================为每层地层创建一个TIN范围==================================

print("\n开始为每层地层创建范围...")

try:
    # 设置工作环境
    arcpy.env.workspace = workspace2
    
    # 泰森多边形要素类路径
    thiessen_polygons = os.path.join(workspace, "thiessen_polygons")
    
    # 确保泰森多边形要素类存在
    if not arcpy.Exists(thiessen_polygons):
        print(f"错误：找不到泰森多边形要素类 {thiessen_polygons}")
    else:
        print(f"找到泰森多边形要素类: {thiessen_polygons}")
        
        # 遍历每个地层
        for layer_id in layer_sequence:
            print(f"\n处理地层 {layer_id}...")
            
            # 输出要素类名称
            output_extent = f"extent_layer_{layer_id}"
            output_extent_path = os.path.join(workspace2, output_extent)
            
            # 如果输出要素类已存在则删除
            if arcpy.Exists(output_extent_path):
                arcpy.Delete_management(output_extent_path)
            
            # 从合并表中查询当前地层厚度不为0的实际钻孔
            borehole_list = []
            query = f"地层编号 = {layer_id} AND 地层厚度 <> 0 AND 钻孔类型 = '实际'"
            
            print(f"查询条件: {query}")
            
            with arcpy.da.SearchCursor(os.path.join(workspace2, merged_table), ["钻孔编号"], query) as cursor:
                for row in cursor:
                    borehole_list.append(row[0])
            
            # 检查是否找到符合条件的钻孔
            if not borehole_list:
                print(f"警告：地层 {layer_id} 没有找到厚度不为0的实际钻孔")
                continue
            
            print(f"找到 {len(borehole_list)} 个符合条件的实际钻孔")
            
            # 构建 SQL 查询语句来选择泰森多边形
            # 注意：需要根据泰森多边形中存储钻孔编号的字段名进行调整
            holes_str = "','".join(borehole_list)
            thiessen_query = f"钻孔编号 IN ('{holes_str}')"
            
            print(f"泰森多边形查询条件: {thiessen_query}")
            
            try:
                # 选择符合条件的泰森多边形
                arcpy.MakeFeatureLayer_management(thiessen_polygons, "thiessen_lyr")
                arcpy.SelectLayerByAttribute_management("thiessen_lyr", "NEW_SELECTION", thiessen_query)
                
                # 获取选择的要素数量
                selected_count = int(arcpy.GetCount_management("thiessen_lyr")[0])
                print(f"选择了 {selected_count} 个泰森多边形")
                
                if selected_count > 0:
                    # 将选择的要素保存为新的要素类
                    arcpy.CopyFeatures_management("thiessen_lyr", output_extent_path)
                    print(f"已创建地层 {layer_id} 的范围要素类: {output_extent}")
                else:
                    print(f"警告：未找到地层 {layer_id} 对应的泰森多边形")
                
            except arcpy.ExecuteError:
                # 如果泰森多边形中的字段名与预期不同，尝试查找正确的字段名
                print("泰森多边形查询出错，尝试确定正确的字段名...")
                
                # 获取泰森多边形的字段列表
                fields = [f.name for f in arcpy.ListFields(thiessen_polygons)]
                print(f"泰森多边形字段列表: {fields}")
                
                # 找到可能存储钻孔编号的字段
                possible_fields = [f for f in fields if "钻孔" in f or "编号" in f or "ID" in f.upper() or "HOLE" in f.upper()]
                
                if possible_fields:
                    print(f"可能的钻孔编号字段: {possible_fields}")
                    for field in possible_fields:
                        try:
                            thiessen_query = f"{field} IN ('{holes_str}')"
                            arcpy.SelectLayerByAttribute_management("thiessen_lyr", "NEW_SELECTION", thiessen_query)
                            
                            selected_count = int(arcpy.GetCount_management("thiessen_lyr")[0])
                            print(f"使用字段 {field} 选择了 {selected_count} 个泰森多边形")
                            
                            if selected_count > 0:
                                arcpy.CopyFeatures_management("thiessen_lyr", output_extent_path)
                                print(f"已创建地层 {layer_id} 的范围要素类: {output_extent}")
                                break
                        except:
                            continue
                
                # 如果无法通过字段查询，尝试通过钻孔中心点创建缓冲区
                if not arcpy.Exists(output_extent_path):
                    print("备选方案: 从实际钻孔点创建范围")
                    
                    # 创建临时钻孔点要素类
                    temp_points = "in_memory\\temp_points"
                    if arcpy.Exists(temp_points):
                        arcpy.Delete_management(temp_points)
                    
                    # 创建临时钻孔点图层
                    arcpy.CreateFeatureclass_management("in_memory", "temp_points", "POINT", spatial_reference=sr)
                    arcpy.AddField_management(temp_points, "钻孔编号", "TEXT")
                    
                    # 插入钻孔点
                    with arcpy.da.InsertCursor(temp_points, ["SHAPE@XY", "钻孔编号"]) as icursor:
                        for hole_id in borehole_list:
                            # 查询钻孔坐标
                            with arcpy.da.SearchCursor(os.path.join(workspace2, merged_table), 
                                                   ["X坐标", "Y坐标"], 
                                                   f"钻孔编号 = '{hole_id}' AND 地层编号 = {layer_id}") as scursor:
                                for row in scursor:
                                    x, y = row
                                    if x is not None and y is not None:
                                        icursor.insertRow([(x, y), hole_id])
                                        break
                    
                    # 创建凸包
                    arcpy.MinimumBoundingGeometry_management(temp_points, output_extent_path, "CONVEX_HULL")
                    print(f"已通过凸包方法创建地层 {layer_id} 的范围要素类: {output_extent}")
                    
                    # 清理临时数据
                    arcpy.Delete_management(temp_points)
        
        print("\n所有地层范围创建完成！")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
finally:
    # 清理临时图层
    if arcpy.Exists("thiessen_lyr"):
        arcpy.Delete_management("thiessen_lyr")


#=======================================两个TIN之间拉伸=================================
print("\n开始在TIN之间拉伸创建多面体...")

try:
    # 设置工作环境
    arcpy.env.overwriteOutput = True
    arcpy.env.workspace = workspace2
    
    # 使用指定的坐标系统
    sr = arcpy.SpatialReference(4548)  # CGCS2000
    
    # 获取扩展名为.gdb的目录
    tin_dir = os.path.join(base_path, 'TIN')
    
    # 检查3D Analyst扩展是否可用
    if arcpy.CheckExtension("3D") == "Available":
        arcpy.CheckOutExtension("3D")
        print("已启用3D Analyst扩展")
    else:
        print("错误：3D Analyst扩展不可用，无法进行TIN间拉伸操作")
        raise Exception("缺少所需的扩展许可证：3D Analyst")
    
    # 处理表土层(0)和第一个地层(1)之间的拉伸
    print(f"\n处理 表土层(0) 和 地层{layer_sequence[0]} 之间的拉伸...")
    
    # 定义TIN路径
    tin_surface = os.path.join(tin_dir, f"TIN_0")
    tin_layer1 = os.path.join(tin_dir, f"TIN_{layer_sequence[0]}")
    
    # 定义拉伸边界（使用第一个地层的范围）
    extent_fc = f"extent_layer_{layer_sequence[0]}"
    
    # 定义输出多面体名称
    output_mp = f"multipatch_surface_layer_{layer_sequence[0]}"
    output_mp_path = os.path.join(workspace2, output_mp)
    
    # 检查输入是否存在
    if arcpy.Exists(tin_surface) and arcpy.Exists(tin_layer1) and arcpy.Exists(extent_fc):
        # 如果输出已存在则删除
        if arcpy.Exists(output_mp_path):
            arcpy.Delete_management(output_mp_path)
            print(f"删除已存在的多面体 {output_mp}")
        
        try:
            print(f"正在创建表土层多面体 {output_mp}...")
            
            # 创建临时图层用于拉伸操作
            arcpy.MakeFeatureLayer_management(extent_fc, "extent_lyr")
            
            # 执行TIN之间的拉伸
            arcpy.ddd.ExtrudeBetween(
                in_tin1=tin_surface,
                in_tin2=tin_layer1,
                in_feature_class="extent_lyr",
                out_feature_class=output_mp_path
            )
            
            # 获取创建的多面体要素数
            mp_count = int(arcpy.GetCount_management(output_mp_path)[0])
            
            print(f"成功创建表土层多面体 {output_mp}")
            print(f"  要素数量: {mp_count}")
            
            # 为多面体添加属性信息
            if mp_count > 0:
                print("添加表土层属性信息...")
                
                # 添加地层信息字段
                arcpy.AddField_management(output_mp_path, "地层编号", "LONG")
                arcpy.AddField_management(output_mp_path, "地层名称", "TEXT", field_length=50)
                arcpy.AddField_management(output_mp_path, "上层编号", "LONG")
                arcpy.AddField_management(output_mp_path, "下层编号", "LONG")
                
                # 更新属性
                with arcpy.da.UpdateCursor(output_mp_path, ["地层编号", "地层名称", "上层编号", "下层编号"]) as cursor:
                    for row in cursor:
                        row[0] = layer_sequence[0]  # 第一个地层编号
                        row[1] = "表土层"  # 表土层的名称
                        row[2] = 0  # 上层为表土层
                        row[3] = layer_sequence[0]  # 下层为第一个地层
                        cursor.updateRow(row)
                
                print("表土层属性更新完成")
                
        except arcpy.ExecuteError:
            print(f"处理表土层和地层 {layer_sequence[0]} 时出错:")
            print(arcpy.GetMessages())
        finally:
            # 清理临时图层
            if arcpy.Exists("extent_lyr"):
                arcpy.Delete_management("extent_lyr")
    else:
        print(f"错误：表土层拉伸所需的输入数据不完整")
    
    # 遍历地层序列
    for i in range(len(layer_sequence) - 1):
        current_layer = layer_sequence[i]
        next_layer = layer_sequence[i + 1]
        
        print(f"\n处理 地层{current_layer} 和 地层{next_layer} 之间的拉伸...")
        
        # 定义TIN路径
        tin_current = os.path.join(tin_dir, f"TIN_{current_layer}")
        tin_next = os.path.join(tin_dir, f"TIN_{next_layer}")
        
        # 定义拉伸边界（使用下一个地层的范围）
        extent_fc = f"extent_layer_{next_layer}"
        
        # 定义输出多面体名称
        output_mp = f"multipatch_layer_{next_layer}"
        output_mp_path = os.path.join(workspace2, output_mp)
        
        # 检查输入是否存在
        if not arcpy.Exists(tin_current):
            print(f"错误：找不到TIN {tin_current}")
            continue
        
        if not arcpy.Exists(tin_next):
            print(f"错误：找不到TIN {tin_next}")
            continue
        
        if not arcpy.Exists(extent_fc):
            print(f"错误：找不到范围要素类 {extent_fc}")
            continue
        
        # 如果输出已存在则删除
        if arcpy.Exists(output_mp_path):
            arcpy.Delete_management(output_mp_path)
            print(f"删除已存在的多面体 {output_mp}")
        
        try:
            print(f"正在创建多面体 {output_mp}...")
            
            # 创建临时图层用于拉伸操作
            arcpy.MakeFeatureLayer_management(extent_fc, "extent_lyr")
            
            # 执行TIN之间的拉伸
            arcpy.ddd.ExtrudeBetween(
                in_tin1=tin_current,
                in_tin2=tin_next,
                in_feature_class="extent_lyr",
                out_feature_class=output_mp_path
            )
            # 获取创建的多面体要素数
            mp_count = int(arcpy.GetCount_management(output_mp_path)[0])
            
            print(f"成功创建多面体 {output_mp}")
            print(f"  要素数量: {mp_count}")
            
            # 为多面体添加属性信息
            if mp_count > 0:
                print("添加地层属性信息...")
                
                # 添加地层信息字段
                arcpy.AddField_management(output_mp_path, "地层编号", "LONG")
                arcpy.AddField_management(output_mp_path, "地层名称", "TEXT", field_length=50)
                arcpy.AddField_management(output_mp_path, "上层编号", "LONG")
                arcpy.AddField_management(output_mp_path, "下层编号", "LONG")
                
                # 更新属性
                with arcpy.da.UpdateCursor(output_mp_path, ["地层编号", "地层名称", "上层编号", "下层编号"]) as cursor:
                    for row in cursor:
                        row[0] = next_layer
                        row[1] = layer_info[next_layer]
                        row[2] = current_layer
                        row[3] = next_layer
                        cursor.updateRow(row)
                
                print("地层属性更新完成")
                
        except arcpy.ExecuteError:
            print(f"处理地层 {current_layer} 和 {next_layer} 时出错:")
            print(arcpy.GetMessages())
        finally:
            # 清理临时图层
            if arcpy.Exists("extent_lyr"):
                arcpy.Delete_management("extent_lyr")
    
    print("\n所有地层拉伸处理完成!")
    
    # 记录生成的多面体数量
    multipatch_count = 0
    for fc in arcpy.ListFeatureClasses("multipatch_*"):
        multipatch_count += 1
    
    print(f"总共生成 {multipatch_count} 个多面体要素类")

except arcpy.ExecuteError as e:
    print("\nArcGIS错误：")
    print(arcpy.GetMessages())
except Exception as e:
    print(f"\n发生错误：{str(e)}")
    import traceback
    print(traceback.format_exc())
finally:
    # 归还许可证
    arcpy.CheckInExtension("3D")
    print("已归还3D Analyst扩展许可")


print("\n三维地质建模过程全部完成！")