# 确保依赖存在，缺失时用当前解释器自动安装（仅源码运行时；打包成 exe 后不执行 pip）
def _ensure_deps():
    import sys
    try:
        import pandas as pd  # noqa: F401
        import openpyxl  # noqa: F401
        return
    except ModuleNotFoundError:
        if getattr(sys, "frozen", False):
            raise  # 打包后的 exe 内缺库直接报错，不尝试 pip
        import subprocess
        pkgs = ['pandas', 'openpyxl']
        print(f"正在安装依赖: {', '.join(pkgs)} ...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-q'] + pkgs)
        print("安装完成，继续执行。\n")


_ensure_deps()
import os
import pandas as pd


class Weather:
    # 自定义路径时，在此相对路径下查找 weather.xlsx（根目录由调用方选择）
    RELATIVE_EXCEL_PATH = os.path.join("RawAssets", "DesignerAssets", "NewDatabase", "logic", "weather.xlsx")

    def __init__(self, branch='stage', custom_excel_path=None):
        # 若指定了自定义 excel 路径，直接使用
        if custom_excel_path:
            self.path = custom_excel_path
            self.current_branch = 'custom'
            self.branches = {}
            return
        # 定义4个分支路径
        self.branches = {
            'stage': 'H:\\zhangjunjie_stage_1\\RawAssets\\DesignerAssets\\NewDatabase\\logic',
            'hotfix': 'H:\\zhangjunjie_obt_hotfix1_1\\RawAssets\\DesignerAssets\\NewDatabase\\logic',
            'review': 'H:\\zhangjunjie_obt_review1_1\\RawAssets\\DesignerAssets\\NewDatabase\\logic',
            'release': 'H:\\zhangjunjie_obt_release3_1\\RawAssets\\DesignerAssets\\NewDatabase\\logic'
        }
        self.current_branch = branch if branch in self.branches else 'stage'
        branch_path = self.branches[self.current_branch]
        self.path = f"{branch_path}\\weather.xlsx"
    
    def read_file(self):
        self.df_weather_type = pd.read_excel(self.path, sheet_name='weatherType',skiprows=4)
        self.df_weather_list = pd.read_excel(self.path, sheet_name='weatherList',skiprows=4)
        return self.df_weather_type, self.df_weather_list

    def get_weather_type(self,weather_id):
        if weather_id in range(5) or weather_id in range(101,107) or weather_id in range(201,214):
            weather_name = self.df_weather_type[self.df_weather_type['id'] == weather_id]['nameDay'].values[0]
        elif weather_id in range(107,119) or weather_id in range(301,306) or weather_id == 399:
            # 对于weather_id在107-118范围内的特殊处理
            if weather_id in range(107,119):
                # 从season_part中提取流星类型信息
                season_part = self.df_weather_type[self.df_weather_type['id'] == weather_id].iloc[0,7]
                
                # 1. 移除季节部分(如"春季"、"夏季"等)
                meteor_type = season_part
                for season in ['春季', '夏季', '秋季', '冬季']:
                    if season in meteor_type:
                        meteor_type = meteor_type.replace(season, '').strip()
                        break
                
                # 2. 确保流星类型中包含"流星雨"关键词
                if '流星雨' not in meteor_type and meteor_type:
                    meteor_type = f"{meteor_type}流星雨"
                
                # 3. 获取地点信息并处理
                weather_part = self.df_weather_type[self.df_weather_type['id'] == weather_id]['nameDay'].values[0]
                # 从地点信息中去掉可能的"流星雨"关键词
                if '-流星雨' in weather_part:
                    weather_part = weather_part.replace('-流星雨', '').strip()
                elif '流星雨' in weather_part:
                    weather_part = weather_part.replace('流星雨', '').strip()
                
                # 4. 清理meteor_type和weather_part中可能的多余连字符
                meteor_type = meteor_type.strip('-')  # 移除前后可能的连字符
                weather_part = weather_part.strip('-')  # 移除前后可能的连字符
                
                # 构建最终的天气名称，格式为"XX流星雨-XX"（如"小规模流星雨-渔村"）
                # 只有当meteor_type和weather_part都不为空时才添加连字符
                if meteor_type and weather_part:
                    weather_name = f"{meteor_type}-{weather_part}"
                elif meteor_type:
                    weather_name = meteor_type
                else:
                    weather_name = weather_part
            else:
                weather_name = f"{self.df_weather_type[self.df_weather_type['id'] == weather_id].iloc[0,7]}-{self.df_weather_type[self.df_weather_type['id'] == weather_id]['nameDay'].values[0]}"
        else:
            # 未在上述区间内的 id（如 119、200+ 等）尝试从表取 nameDay，否则返回未知（Excel 可能读成浮点）
            _id = int(weather_id) if isinstance(weather_id, float) else weather_id
            match = self.df_weather_type[self.df_weather_type['id'] == _id]
            if not match.empty and 'nameDay' in match.columns:
                weather_name = match['nameDay'].values[0]
            else:
                weather_name = f"未知({weather_id})"
        return weather_name
        
    # 输出格式：每行宽度（用于对齐与分隔线）
    _OUTPUT_WIDTH = 44

    def _format_day_header(self, month, day):
        """返回单日标题的若干行（分隔线 + 标题 + 下划线）"""
        title = f"  ◆ {month}月{day}日 天气情况"
        return [
            "═" * self._OUTPUT_WIDTH,
            title,
            "─" * self._OUTPUT_WIDTH,
        ]

    # 特殊天气 ID 集合：104、106、107-121、211-213、301-305（用于单日「本日特殊天气」展示）
    SPECIAL_WEATHER_IDS = {104, 106} | set(range(107, 122)) | set(range(211, 214)) | set(range(301, 306))

    def get_special_weather_for_day(self, month, day):
        """获取指定日期的特殊天气时段（仅 ID 在 SPECIAL_WEATHER_IDS 内），合并连续相同 ID。
        有则返回格式化字符串，无则返回空字符串（不显示该日）。"""
        day_data = self.df_weather_list[(self.df_weather_list["month"] == month) & (self.df_weather_list["day"] == day)]
        if day_data.empty:
            return ""
        row = day_data.iloc[0]
        lines = []
        i = 0
        while i < 24:
            w_id = self._cell_to_id(row[f"h{i}"])
            if w_id is None or w_id not in self.SPECIAL_WEATHER_IDS:
                i += 1
                continue
            w_name = self.get_weather_type(row[f"h{i}"])
            start = i
            while i + 1 < 24 and self._cell_to_id(row[f"h{i+1}"]) == w_id:
                i += 1
            end = i
            end_display = 24 if end == 23 else end
            id_suffix = f" ({w_id})" if w_id is not None else ""
            if start == end:
                if start == 23:
                    lines.append(f"    · {start}~24点  {w_name}{id_suffix}")
                else:
                    lines.append(f"    · {start}点  {w_name}{id_suffix}")
            else:
                lines.append(f"    · {start}~{end_display}点  {w_name}{id_suffix}")
            i += 1
        return "\n".join(lines) if lines else ""

    def get_special_weather_for_range(self, start_month, start_day, end_month, end_day):
        """获取日期范围内每日的特殊天气，按日显示并带具体时间段；无特殊天气的日期不显示。"""
        parts = []
        for index, row in self.df_weather_list.iterrows():
            current_month = int(row["month"])
            current_day = int(row["day"])
            in_range = False
            if start_month == end_month:
                in_range = start_month == current_month and start_day <= current_day <= end_day
            else:
                if current_month == start_month and current_day >= start_day:
                    in_range = True
                elif start_month < current_month < end_month:
                    in_range = True
                elif current_month == end_month and current_day <= end_day:
                    in_range = True
            if not in_range:
                continue
            day_special = self.get_special_weather_for_day(current_month, current_day)
            if not day_special:
                continue
            # 统一块宽，与单日展示一致
            header = f"  ┌─ {current_month}月{current_day}日"
            parts.append(header)
            parts.append(day_special)
            parts.append("  └" + "─" * 10)
            parts.append("")
        return "\n".join(parts).strip() if parts else "该范围内无特殊天气"

    def _format_hourly_weather_table(self, row):
        """单天逐段表格行：[(时间段, 天气名, ID), ...]，用于 GUI 表格展示。"""
        table_rows = []
        i = 0
        while i < 24:
            w_id = self._cell_to_id(row[f'h{i}'])
            w_name = self.get_weather_type(row[f'h{i}'])
            start = i
            while i + 1 < 24 and self._cell_to_id(row[f'h{i+1}']) == w_id:
                i += 1
            end = i
            end_display = 24 if end == 23 else end
            if start == end:
                time_str = f"{start}~24点" if start == 23 else f"{start}点"
            else:
                time_str = f"{start}~{end_display}点"
            table_rows.append((time_str, w_name, str(w_id) if w_id is not None else ""))
            i += 1
        return table_rows

    def _format_hourly_weather(self, row):
        """格式化单天每小时天气数据；连续相同天气合并为「起始~结束点：天气名」"""
        weather_ids = []
        for i in range(24):
            weather_ids.append(row[f'h{i}'])
        # 合并连续相同天气的小时
        hourly_data = []
        i = 0
        while i < 24:
            w_id = self._cell_to_id(row[f'h{i}'])
            w_name = self.get_weather_type(row[f'h{i}'])
            start = i
            while i + 1 < 24 and self._cell_to_id(row[f'h{i+1}']) == w_id:
                i += 1
            end = i
            # 最后一小时（23点）显示为 23~24点，与「到24点」一致；结果中附带天气 ID
            end_display = 24 if end == 23 else end
            id_suffix = f" ({w_id})" if w_id is not None else ""
            if start == end:
                if start == 23:
                    hourly_data.append(f"    {start}~24点：{w_name}{id_suffix}")
                else:
                    hourly_data.append(f"    {start}点：{w_name}{id_suffix}")
            else:
                hourly_data.append(f"    {start}~{end_display}点：{w_name}{id_suffix}")
            i += 1
        return weather_ids, hourly_data
    
    def _save_to_file(self, data):
        """将天气数据保存到txt文件"""
        import os
        from datetime import datetime
        
        # 创建输出目录，放到当前Python代码文件所在目录
        current_script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(current_script_dir, 'output')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 生成文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_path = os.path.join(output_dir, f'weather_all_{timestamp}.txt')
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(data)
        
        return file_path
    
    def get_weather_list_by_day(self, month=None, day=None, show_all=False, save_to_file=False, 
                               start_month=None, start_day=None, end_month=None, end_day=None):
        """
        获取指定日期、日期范围或所有日期的天气列表
        
        :param month: 月份（可选）
        :param day: 日期（可选）
        :param show_all: 是否显示所有日期数据
        :param save_to_file: 是否保存到txt文件
        :param start_month: 开始月份（可选，用于日期范围）
        :param start_day: 开始日期（可选，用于日期范围）
        :param end_month: 结束月份（可选，用于日期范围）
        :param end_day: 结束日期（可选，用于日期范围）
        :return: (weather_data, weather_data_translate, output_file_path, table_columns, table_rows)
           table_columns/table_rows 为 None 或空时表示无表格数据，仅用文本展示。
        """
        weather_data = []
        weather_data_translate = ""
        output_file_path = None
        
        table_columns = None
        table_rows = None

        # 处理指定日期的情况
        if month and day:
            # 获取指定日期的数据
            day_data = self.df_weather_list[(self.df_weather_list['month'] == month) & (self.df_weather_list['day'] == day)]
            if not day_data.empty:
                row = day_data.iloc[0]
                daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                weather_data = daily_weather_ids
                header_lines = self._format_day_header(month, day)
                weather_data_translate = "\n".join(header_lines) + "\n" + "\n".join(hourly_data)
                table_columns = ["时间段", "天气", "ID"]
                table_rows = self._format_hourly_weather_table(row)
        
        # 处理日期范围的情况
        elif start_month and start_day and end_month and end_day:
            # 准备存储范围内的天气数据
            range_weather_translate = []
            range_weather_data = []
            
            # 遍历所有日期数据
            for index, row in self.df_weather_list.iterrows():
                current_month = int(row['month'])
                current_day = int(row['day'])
                
                # 检查当前日期是否在范围内
                # 同月份的情况
                if start_month == end_month:
                    if start_month == current_month and start_day <= current_day <= end_day:
                        # 在范围内，添加数据
                        range_weather_translate.extend(self._format_day_header(current_month, current_day))
                        daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                        range_weather_data.extend(daily_weather_ids)
                        range_weather_translate.extend(hourly_data)
                        range_weather_translate.append("")
                # 跨月份的情况
                else:
                    # 开始月份
                    if current_month == start_month and current_day >= start_day:
                        range_weather_translate.extend(self._format_day_header(current_month, current_day))
                        daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                        range_weather_data.extend(daily_weather_ids)
                        range_weather_translate.extend(hourly_data)
                        range_weather_translate.append("")
                    # 中间月份
                    elif start_month < current_month < end_month:
                        range_weather_translate.extend(self._format_day_header(current_month, current_day))
                        daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                        range_weather_data.extend(daily_weather_ids)
                        range_weather_translate.extend(hourly_data)
                        range_weather_translate.append("")
                    # 结束月份
                    elif current_month == end_month and current_day <= end_day:
                        range_weather_translate.extend(self._format_day_header(current_month, current_day))
                        daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                        range_weather_data.extend(daily_weather_ids)
                        range_weather_translate.extend(hourly_data)
                        range_weather_translate.append("")
            
            # 生成结果
            if range_weather_translate:
                weather_data = range_weather_data
                weather_data_translate = "\n".join(range_weather_translate)
                table_columns = ["日期", "时间段", "天气", "ID"]
                table_rows = []
                for index, row in self.df_weather_list.iterrows():
                    current_month = int(row['month'])
                    current_day = int(row['day'])
                    if start_month == end_month:
                        if start_month != current_month or not (start_day <= current_day <= end_day):
                            continue
                    else:
                        if current_month == start_month and current_day < start_day:
                            continue
                        if current_month == end_month and current_day > end_day:
                            continue
                        if current_month < start_month or current_month > end_month:
                            continue
                    date_str = f"{current_month}月{current_day}日"
                    for time_str, w_name, id_str in self._format_hourly_weather_table(row):
                        table_rows.append((date_str, time_str, w_name, id_str))
                
                # 如果需要保存到文件
                if save_to_file:
                    output_file_path = self._save_to_file(weather_data_translate)
        
        # 处理显示所有日期的情况
        elif show_all:
            # 遍历所有日期数据
            all_weather_translate = []
            all_weather_data = []
            
            for index, row in self.df_weather_list.iterrows():
                month_val = int(row['month'])
                day_val = int(row['day'])
                all_weather_translate.extend(self._format_day_header(month_val, day_val))
                daily_weather_ids, hourly_data = self._format_hourly_weather(row)
                all_weather_data.extend(daily_weather_ids)
                all_weather_translate.extend(hourly_data)
                all_weather_translate.append("")
            
            weather_data = all_weather_data
            weather_data_translate = "\n".join(all_weather_translate)
            table_columns = ["日期", "时间段", "天气", "ID"]
            table_rows = []
            for index, row in self.df_weather_list.iterrows():
                month_val = int(row['month'])
                day_val = int(row['day'])
                date_str = f"{month_val}月{day_val}日"
                for time_str, w_name, id_str in self._format_hourly_weather_table(row):
                    table_rows.append((date_str, time_str, w_name, id_str))
            
            # 如果需要保存到文件
            if save_to_file:
                output_file_path = self._save_to_file(weather_data_translate)
        
        return weather_data, weather_data_translate, output_file_path, table_columns, table_rows
            
    def find_weather_id(self,weather_id):
        """
        查找包含指定weather_id的所有日期和时间
        :param weather_id: 要查找的天气ID
        :return: 格式化的字符串，显示包含该天气ID的所有日期和时间，日期之间进行换行
        """
        # 存储结果的字典，键为日期字符串，值为该日期中包含指定weather_id的小时列表
        result_dict = {}
        
        # 遍历weatherList的每一行数据
        for index, row in self.df_weather_list.iterrows():
            month = int(row['month'])
            day = int(row['day'])
            date_str = f"{month}月{day}日"
            
            # 检查每小时的天气ID
            matching_hours = []
            for i in range(24):
                if f'h{i}' in row and row[f'h{i}'] == weather_id:
                    matching_hours.append(f"{i}点")
            
            # 如果该日期有匹配的小时，添加到结果字典
            if matching_hours:
                result_dict[date_str] = matching_hours
        
        # 构建输出字符串
        if not result_dict:
            return f"未找到weather_id为{weather_id}的天气数据"
        
        output_parts = []
        for date_str, hours in result_dict.items():
            # 格式化小时列表，如："18点，19点，20点"
            hours_str = "，".join(hours)
            output_parts.append(f"{date_str}：{hours_str}")
        
        # 组合最终输出，使日期之间进行换行
        formatted_output = "\n".join(output_parts)
        final_output = f"以下日期存在对应id天气：\n{formatted_output}"
        return final_output
    
    def get_special_weather_in_range(self, start_month, start_day, end_month, end_day, save_to_file=False):
        """
        获取指定日期范围内的特殊天气时间
        特殊天气：天气id不为101-106，201-204的天气均为特殊天气
        
        :param start_month: 开始月份
        :param start_day: 开始日期
        :param end_month: 结束月份
        :param end_day: 结束日期
        :param save_to_file: 是否保存到txt文件
        :return: (special_weather_list, formatted_output, output_file_path)
        """
        # 定义正常天气ID范围
        normal_weather_ranges = [(101, 106), (201, 204)]
        
        def is_special_weather(weather_id):
            """检查是否为特殊天气"""
            for start, end in normal_weather_ranges:
                if start <= weather_id <= end:
                    return False
            return True
        
        # 存储特殊天气数据
        special_weather_list = []
        formatted_output = ""
        output_file_path = None
        
        # 遍历所有日期数据
        for index, row in self.df_weather_list.iterrows():
            current_month = int(row['month'])
            current_day = int(row['day'])
            
            # 检查当前日期是否在范围内
            is_in_range = False
            # 同月份的情况
            if start_month == end_month:
                if start_month == current_month and start_day <= current_day <= end_day:
                    is_in_range = True
            # 跨月份的情况
            else:
                if (start_month == current_month and current_day >= start_day) or \
                   (start_month < current_month < end_month) or \
                   (end_month == current_month and current_day <= end_day):
                    is_in_range = True
            
            if is_in_range:
                # 检查每小时的天气
                for i in range(24):
                    weather_id = row[f'h{i}']
                    if is_special_weather(weather_id):
                        weather_name = self.get_weather_type(weather_id)
                        
                        # 格式化输出：例如"  1月 1日  12点～13点  彩虹"
                        special_weather_list.append([current_month, current_day, i, i+1, weather_name])
        
        # 生成格式化输出
        if special_weather_list:
            formatted_parts = []
            for month, day, start_hour, end_hour, weather_name in special_weather_list:
                formatted_parts.append(f"  {month}月{day:>2}日  {start_hour:>2}点～{end_hour:>2}点  {weather_name}")
            formatted_output = "\n".join(formatted_parts)
            
            # 如果需要保存到文件
            if save_to_file:
                output_file_path = self._save_to_file(formatted_output)
        
        return special_weather_list, formatted_output, output_file_path
    
    @staticmethod
    def _cell_to_id(x):
        """将 Excel 单元格值统一转为整数 ID，便于与 weather_ids 比较（避免 303.0、'303' 等漏匹配）"""
        if x is None or (hasattr(pd, 'isna') and pd.isna(x)):
            return None
        try:
            return int(float(x))
        except (TypeError, ValueError):
            return None

    def find_weather_ids_time_ranges(self, weather_ids, save_to_file=False):
        """
        查找指定weather_ids的天气有哪几天的几点到几点。
        输出按查询的 ID 顺序分组，同一 ID 内按时间（月、日、起始小时）排序。
        :param weather_ids: 要查找的天气ID列表，如[110, 121, 301]
        :param save_to_file: 是否保存到txt文件
        :return: (weather_ranges, formatted_output, output_file_path, table_columns, table_rows)
        """
        weather_ids_set = set(int(x) for x in weather_ids)
        weather_ranges = []  # (month, day, start_hour, end_hour, weather_id_int, weather_name)

        for index, row in self.df_weather_list.iterrows():
            current_month = int(row['month'])
            current_day = int(row['day'])
            current_ranges = []
            start_hour = None
            start_id = None  # 当前连续段的天气 ID，ID 变化时必须结束当前段、另起一段

            for i in range(24):
                w_id = self._cell_to_id(row[f'h{i}'])
                if w_id is not None and w_id in weather_ids_set:
                    if start_hour is None:
                        start_hour = i
                        start_id = int(w_id)
                    elif int(w_id) != start_id:
                        # 同一行内 ID 变化：先结束当前段，再开始新段
                        end_hour = i
                        w_name = self.get_weather_type(row[f'h{start_hour}'])
                        current_ranges.append([current_month, current_day, start_hour, end_hour, start_id, w_name])
                        start_hour = i
                        start_id = int(w_id)
                else:
                    if start_hour is not None:
                        end_hour = i
                        w_name = self.get_weather_type(row[f'h{start_hour}'])
                        current_ranges.append([current_month, current_day, start_hour, end_hour, start_id, w_name])
                        start_hour = None
                        start_id = None

            if start_hour is not None:
                w_name = self.get_weather_type(row[f'h{start_hour}'])
                current_ranges.append([current_month, current_day, start_hour, 24, start_id, w_name])
            weather_ranges.extend(current_ranges)
        
        # 按查询的 ID 顺序分组，组内按 (月, 日, 起始小时) 排序（key 统一为 int 避免 109 vs 109.0 导致漏显）
        formatted_output = ""
        output_file_path = None
        
        table_columns = ["日期", "时间段", "天气"]
        table_rows = []

        if not weather_ranges:
            formatted_output = f"未找到weather_id为{', '.join(map(str, weather_ids))}的天气数据"
            return weather_ranges, formatted_output, output_file_path, table_columns, table_rows
        
        # 按查询顺序输出每个 ID：直接从 weather_ranges 中筛出该 ID 的时段，避免 dict key 类型差异导致漏显
        def segment_id_as_int(r):
            try:
                v = r[4]
                return int(v) if v is not None else None
            except (TypeError, ValueError):
                return None

        parts = []
        weather_ids_int = [int(x) for x in weather_ids]
        for wid in weather_ids_int:
            ranges_list = []
            for r in weather_ranges:
                if segment_id_as_int(r) == wid:
                    month, day, sh, eh, _, w_name = r
                    ranges_list.append((month, day, sh, eh, w_name))
            ranges_list.sort(key=lambda x: (x[0], x[1], x[2]))  # 月、日、起始小时
            w_name_header = self.get_weather_type(wid) if ranges_list else ""
            parts.append(f"--- 天气 ID {wid} {w_name_header} ---")
            if ranges_list:
                for month, day, start_hour, end_hour, w_name in ranges_list:
                    parts.append(f"  {month}月{day:>2}日  {start_hour:>2}点～{end_hour:>2}点  {w_name}")
                    time_str = f"{start_hour}~{end_hour}点" if start_hour != end_hour else f"{start_hour}点"
                    table_rows.append((f"{month}月{day}日", time_str, w_name))
            else:
                parts.append("  （无）")
            parts.append("")
        
        formatted_output = "\n".join(parts).strip()
        if save_to_file:
            output_file_path = self._save_to_file(formatted_output)
        
        return weather_ranges, formatted_output, output_file_path, table_columns, table_rows

    @staticmethod
    def compare_branches(branch_a, branch_b, save_to_file=False):
        """
        对比两个分支路径下的 weather.xlsx，返回差别说明。

        :param branch_a: 分支名，如 'stage'、'hotfix'、'review'、'release'
        :param branch_b: 另一分支名
        :param save_to_file: 是否将对比结果保存为 txt
        :return: (diff_dict, formatted_report, output_file_path)
                 diff_dict 含 type_only_a, type_only_b, type_value_diff, list_only_a, list_only_b, list_hour_diff 等
        """
        wa = Weather(branch=branch_a)
        wb = Weather(branch=branch_b)
        if not os.path.isfile(wa.path):
            raise FileNotFoundError(f"分支 {branch_a} 文件不存在: {wa.path}")
        if not os.path.isfile(wb.path):
            raise FileNotFoundError(f"分支 {branch_b} 文件不存在: {wb.path}")

        wa.read_file()
        wb.read_file()

        def _norm_id(x):
            if pd.isna(x):
                return None
            return int(x) if isinstance(x, float) else x

        # ---------- weatherType 对比（以 id 为键）----------
        da = wa.df_weather_type.copy()
        db = wb.df_weather_type.copy()
        da['_id'] = da['id'].map(_norm_id)
        db['_id'] = db['id'].map(_norm_id)
        ids_a = set(da['_id'].dropna().unique())
        ids_b = set(db['_id'].dropna().unique())
        type_only_a = sorted(ids_a - ids_b)
        type_only_b = sorted(ids_b - ids_a)
        common_ids = sorted(ids_a & ids_b)

        # 选共同列（排除 _id）比较
        cols_type = [c for c in da.columns if c in db.columns and c != '_id' and c != 'id']
        type_value_diff = []  # [(id, col, val_a, val_b), ...]
        for wid in common_ids:
            row_a = da[da['_id'] == wid].iloc[0]
            row_b = db[db['_id'] == wid].iloc[0]
            for col in cols_type:
                v_a, v_b = row_a[col], row_b[col]
                if pd.isna(v_a):
                    v_a = None
                if pd.isna(v_b):
                    v_b = None
                if v_a != v_b:
                    type_value_diff.append((wid, col, v_a, v_b))

        # ---------- weatherList 对比（以 month+day 为键）----------
        la = wa.df_weather_list.copy()
        lb = wb.df_weather_list.copy()
        la['_key'] = la['month'].astype(int).astype(str) + '-' + la['day'].astype(int).astype(str)
        lb['_key'] = lb['month'].astype(int).astype(str) + '-' + lb['day'].astype(int).astype(str)
        keys_a = set(la['_key'].unique())
        keys_b = set(lb['_key'].unique())
        list_only_a = sorted(keys_a - keys_b, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
        list_only_b = sorted(keys_b - keys_a, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
        common_keys = sorted(keys_a & keys_b, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))

        hour_cols = [f'h{i}' for i in range(24)]
        list_hour_diff = []  # [(month, day, hour, id_a, id_b), ...]
        for key in common_keys:
            row_a = la[la['_key'] == key].iloc[0]
            row_b = lb[lb['_key'] == key].iloc[0]
            month, day = int(row_a['month']), int(row_a['day'])
            for i in range(24):
                col = f'h{i}'
                id_a = row_a[col]
                id_b = row_b[col]
                if _norm_id(id_a) != _norm_id(id_b):
                    list_hour_diff.append((month, day, i, id_a, id_b))

        diff_dict = {
            'branch_a': branch_a,
            'branch_b': branch_b,
            'path_a': wa.path,
            'path_b': wb.path,
            'type_only_a': type_only_a,
            'type_only_b': type_only_b,
            'type_value_diff': type_value_diff,
            'list_only_a': list_only_a,
            'list_only_b': list_only_b,
            'list_hour_diff': list_hour_diff,
        }

        # 生成可读报告
        lines = [
            "═" * 60,
            f"  分支对比：{branch_a}  vs  {branch_b}",
            "═" * 60,
            "",
            f"  路径 A: {wa.path}",
            f"  路径 B: {wb.path}",
            "",
            "─" * 60,
            "  【weatherType】天气类型表",
            "─" * 60,
        ]
        if type_only_a:
            lines.append(f"  仅存在于 {branch_a} 的 id: {type_only_a}")
        else:
            lines.append(f"  仅存在于 {branch_a} 的 id: 无")
        if type_only_b:
            lines.append(f"  仅存在于 {branch_b} 的 id: {type_only_b}")
        else:
            lines.append(f"  仅存在于 {branch_b} 的 id: 无")
        if type_value_diff:
            lines.append(f"  同 id 下字段取值不同（共 {len(type_value_diff)} 处）:")
            for wid, col, v_a, v_b in type_value_diff:
                lines.append(f"    id={wid}, 列「{col}」: A={v_a!r}  →  B={v_b!r}")
        else:
            lines.append("  同 id 下字段取值不同: 无")
        lines.append("")
        lines.append("─" * 60)
        lines.append("  【weatherList】每日天气表")
        lines.append("─" * 60)
        if list_only_a:
            lines.append(f"  仅存在于 {branch_a} 的日期（共 {len(list_only_a)} 天）:")
            for k in list_only_a:
                lines.append(f"    {k}")
        else:
            lines.append(f"  仅存在于 {branch_a} 的日期: 无")
        if list_only_b:
            lines.append(f"  仅存在于 {branch_b} 的日期（共 {len(list_only_b)} 天）:")
            for k in list_only_b:
                lines.append(f"    {k}")
        else:
            lines.append(f"  仅存在于 {branch_b} 的日期: 无")
        if list_hour_diff:
            lines.append(f"  同一日期下小时天气不同（共 {len(list_hour_diff)} 处）:")
            for month, day, hour, id_a, id_b in list_hour_diff:
                lines.append(f"    {month}月{day}日 {hour}点: A={id_a}  →  B={id_b}")
        else:
            lines.append("  同一日期下小时天气不同: 无")
        lines.append("")
        lines.append("═" * 60)

        formatted_report = "\n".join(lines)
        output_file_path = None
        if save_to_file:
            output_file_path = Weather._save_to_file_static(
                formatted_report,
                f"weather_compare_{branch_a}_vs_{branch_b}"
            )

        return diff_dict, formatted_report, output_file_path

    @staticmethod
    def compare_two_paths(excel_path_a, excel_path_b, label_a=None, label_b=None, save_to_file=False):
        """
        对比两个 weather.xlsx 文件路径的差异（选择原理同加载：传入的为项目根目录或直接传 excel 完整路径）。
        若传入的是项目根目录，则自动拼接 RELATIVE_EXCEL_PATH 得到 weather.xlsx。

        :param excel_path_a: 路径 A（项目根目录或 weather.xlsx 的完整路径）
        :param excel_path_b: 路径 B
        :param label_a: 报告中路径 A 的显示名，默认用 excel_path_a
        :param label_b: 报告中路径 B 的显示名，默认用 excel_path_b
        :param save_to_file: 是否保存对比结果到 txt
        :return: (diff_dict, formatted_report, output_file_path)
        """
        def _to_excel_path(p):
            p = p.strip()
            if not p:
                return None
            if p.endswith('.xlsx') and os.path.isfile(p):
                return p
            full = os.path.join(p, Weather.RELATIVE_EXCEL_PATH)
            if os.path.isfile(full):
                return full
            if os.path.isfile(p):
                return p
            return full  # 让后续 isfile 报错

        path_a = _to_excel_path(excel_path_a)
        path_b = _to_excel_path(excel_path_b)
        if not path_a or not os.path.isfile(path_a):
            raise FileNotFoundError(f"路径 A 对应的 weather.xlsx 不存在: {path_a or excel_path_a}")
        if not path_b or not os.path.isfile(path_b):
            raise FileNotFoundError(f"路径 B 对应的 weather.xlsx 不存在: {path_b or excel_path_b}")

        wa = Weather(custom_excel_path=path_a)
        wb = Weather(custom_excel_path=path_b)
        wa.read_file()
        wb.read_file()

        name_a = label_a if label_a is not None else path_a
        name_b = label_b if label_b is not None else path_b

        def _norm_id(x):
            if pd.isna(x):
                return None
            return int(x) if isinstance(x, float) else x

        da = wa.df_weather_type.copy()
        db = wb.df_weather_type.copy()
        da['_id'] = da['id'].map(_norm_id)
        db['_id'] = db['id'].map(_norm_id)
        ids_a = set(da['_id'].dropna().unique())
        ids_b = set(db['_id'].dropna().unique())
        type_only_a = sorted(ids_a - ids_b)
        type_only_b = sorted(ids_b - ids_a)
        common_ids = sorted(ids_a & ids_b)
        cols_type = [c for c in da.columns if c in db.columns and c != '_id' and c != 'id']
        type_value_diff = []
        for wid in common_ids:
            row_a = da[da['_id'] == wid].iloc[0]
            row_b = db[db['_id'] == wid].iloc[0]
            for col in cols_type:
                v_a, v_b = row_a[col], row_b[col]
                if pd.isna(v_a):
                    v_a = None
                if pd.isna(v_b):
                    v_b = None
                if v_a != v_b:
                    type_value_diff.append((wid, col, v_a, v_b))

        la = wa.df_weather_list.copy()
        lb = wb.df_weather_list.copy()
        la['_key'] = la['month'].astype(int).astype(str) + '-' + la['day'].astype(int).astype(str)
        lb['_key'] = lb['month'].astype(int).astype(str) + '-' + lb['day'].astype(int).astype(str)
        keys_a = set(la['_key'].unique())
        keys_b = set(lb['_key'].unique())
        list_only_a = sorted(keys_a - keys_b, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
        list_only_b = sorted(keys_b - keys_a, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
        common_keys = sorted(keys_a & keys_b, key=lambda x: (int(x.split('-')[0]), int(x.split('-')[1])))
        list_hour_diff = []
        for key in common_keys:
            row_a = la[la['_key'] == key].iloc[0]
            row_b = lb[lb['_key'] == key].iloc[0]
            month, day = int(row_a['month']), int(row_a['day'])
            for i in range(24):
                col = f'h{i}'
                id_a, id_b = row_a[col], row_b[col]
                if _norm_id(id_a) != _norm_id(id_b):
                    list_hour_diff.append((month, day, i, id_a, id_b))

        diff_dict = {
            'path_a': path_a,
            'path_b': path_b,
            'type_only_a': type_only_a,
            'type_only_b': type_only_b,
            'type_value_diff': type_value_diff,
            'list_only_a': list_only_a,
            'list_only_b': list_only_b,
            'list_hour_diff': list_hour_diff,
        }

        lines = [
            "═" * 60,
            f"  分支对比：路径 A  vs  路径 B",
            "═" * 60,
            "",
            f"  路径 A: {name_a}",
            f"  路径 B: {name_b}",
            "",
            f"  文件 A: {path_a}",
            f"  文件 B: {path_b}",
            "",
            "─" * 60,
            "  【weatherType】天气类型表",
            "─" * 60,
        ]
        if type_only_a:
            lines.append(f"  仅存在于路径 A 的 id: {type_only_a}")
        else:
            lines.append(f"  仅存在于路径 A 的 id: 无")
        if type_only_b:
            lines.append(f"  仅存在于路径 B 的 id: {type_only_b}")
        else:
            lines.append(f"  仅存在于路径 B 的 id: 无")
        if type_value_diff:
            lines.append(f"  同 id 下字段取值不同（共 {len(type_value_diff)} 处）:")
            for wid, col, v_a, v_b in type_value_diff:
                lines.append(f"    id={wid}, 列「{col}」: A={v_a!r}  →  B={v_b!r}")
        else:
            lines.append("  同 id 下字段取值不同: 无")
        lines.append("")
        lines.append("─" * 60)
        lines.append("  【weatherList】每日天气表")
        lines.append("─" * 60)
        if list_only_a:
            lines.append(f"  仅存在于路径 A 的日期（共 {len(list_only_a)} 天）:")
            for k in list_only_a:
                lines.append(f"    {k}")
        else:
            lines.append(f"  仅存在于路径 A 的日期: 无")
        if list_only_b:
            lines.append(f"  仅存在于路径 B 的日期（共 {len(list_only_b)} 天）:")
            for k in list_only_b:
                lines.append(f"    {k}")
        else:
            lines.append(f"  仅存在于路径 B 的日期: 无")
        if list_hour_diff:
            lines.append(f"  同一日期下小时天气不同（共 {len(list_hour_diff)} 处）:")
            for month, day, hour, id_a, id_b in list_hour_diff:
                lines.append(f"    {month}月{day}日 {hour}点: A={id_a}  →  B={id_b}")
        else:
            lines.append("  同一日期下小时天气不同: 无")
        lines.append("")
        lines.append("═" * 60)

        formatted_report = "\n".join(lines)
        output_file_path = None
        if save_to_file:
            output_file_path = Weather._save_to_file_static(formatted_report, "weather_compare_two_paths")
        return diff_dict, formatted_report, output_file_path

    @staticmethod
    def _save_to_file_static(content, base_name):
        """静态方法：将字符串写入 output 目录，文件名 base_name_时间戳.txt"""
        from datetime import datetime
        current_script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(current_script_dir, 'output')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        file_path = os.path.join(output_dir, f'{base_name}_{timestamp}.txt')
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return file_path


if __name__ == '__main__':
    import os

    # ========== 1. 构造函数用法 ==========
    print("=" * 50)
    print("1. 构造函数 Weather(branch='stage')")
    print("=" * 50)
    weather = Weather(branch='stage')  # 可选: 'stage' | 'hotfix' | 'review' | 'release'
    print(f"当前分支: {weather.current_branch}, 数据路径: {weather.path}\n")

    # ========== 2. read_file() 用法（必须最先执行，后续方法依赖此数据）==========
    print("=" * 50)
    print("2. read_file() - 读取 weather.xlsx")
    print("=" * 50)
    df_type, df_list = weather.read_file()
    print(f"weatherType 行数: {len(df_type)}, weatherList 行数: {len(df_list)}\n")

    # # ========== 3. get_weather_type(weather_id) 用法 ==========
    # print("=" * 50)
    # print("2. read_file() - 读取 weather.xlsx")
    # print("=" * 50)
    # df_type, df_list = weather.read_file()
    # print(f"weatherType 行数: {len(df_type)}, weatherList 行数: {len(df_list)}\n")

    # # ========== 3. get_weather_type(weather_id) 用法 ==========
    # print("=" * 50)
    # print("3. get_weather_type(weather_id) - 根据天气ID获取名称")
    # print("=" * 50)
    # for wid in [101, 102, 119]:
    #     name = weather.get_weather_type(wid)
    #     print(f"  weather_id={wid} -> {name}")
    # print()

    # # ========== 4. get_weather_list_by_day() 用法 ==========
    # # 4.1 指定单日
    # print("=" * 50)
    # print("4. get_weather_list_by_day() - 按日/范围/全部查询")
    # print("=" * 50)
    # print("4.1 指定单日 (month=1, day=1)")
    # weather_data, weather_text, _ = weather.get_weather_list_by_day(month=1, day=1)
    # print(weather_text[:300] + "..." if len(weather_text) > 300 else weather_text)
    # print()

    # 4.2 指定日期范围
    print("4.2 指定日期范围 (1月8日~1月22日, save_to_file=True)")
    _, range_text, range_file, _, _ = weather.get_weather_list_by_day(
        start_month=3, start_day=5, end_month=6, end_day=2, save_to_file=True)
    if range_file:
        print(f"  已保存到: {range_file}")
    print()

    # # 4.3 显示全部日期（仅打印前几行示意）
    # print("4.3 显示全部日期 (show_all=True, 仅打印前500字)")
    # _, all_text, _ = weather.get_weather_list_by_day(show_all=True)
    # print(all_text[:500] + "..." if len(all_text) > 500 else all_text)
    # print()

    # # ========== 5. find_weather_id(weather_id) 用法 ==========
    # print("=" * 50)
    # print("5. find_weather_id(weather_id) - 查找某天气ID出现的日期与小时")
    # print("=" * 50)
    # result = weather.find_weather_id(119)
    # print(result[:400] + "..." if len(result) > 400 else result)
    # print()

    # # ========== 6. get_special_weather_in_range() 用法 ==========
    # print("=" * 50)
    # print("6. get_special_weather_in_range() - 范围内特殊天气")
    # print("=" * 50)
    # special_list, special_text, special_file = weather.get_special_weather_in_range(
    #     start_month=1, start_day=8, end_month=3, end_day=18, save_to_file=True)
    # print(f"  共 {len(special_list)} 条特殊天气")
    # print(special_text[:400] + "..." if len(special_text) > 400 else special_text)
    # if special_file:
    #     print(f"  已保存到: {special_file}")
    # print()

    # # ========== 7. find_weather_ids_time_ranges() 用法 ==========
    # print("=" * 50)
    # print("7. find_weather_ids_time_ranges(weather_ids) - 多ID时间段")
    # print("=" * 50)
    # weather_ids = [119, 120, 121]
    # ranges, ranges_text, ranges_file = weather.find_weather_ids_time_ranges(
    #     weather_ids=weather_ids, save_to_file=True)
    # print(f"  weather_ids={weather_ids}, 共 {len(ranges)} 个时间段")
    # print(ranges_text[:400] + "..." if len(ranges_text) > 400 else ranges_text)
    # if ranges_file:
    #     print(f"  已保存到: {ranges_file}")
    # print()

    # # ========== 检查 output 目录 ==========
    # print("=" * 50)
    # print("output 文件夹中的文件")
    # print("=" * 50)
    # output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    # if os.path.exists(output_dir):
    #     for f in os.listdir(output_dir):
    #         print(f"  - {f}")
    # else:
    #     print("  (目录不存在，未使用 save_to_file=True 时可能为空)")