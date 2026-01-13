import json
import os
from datetime import datetime
from collections import defaultdict, Counter

class GenshinWishAnalyzer:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.wishes = []
        self.uid = ""
        # 映射表：Gacha Type 到可读名称
        self.banner_map = {
            "301": "角色活动祈愿",
            "400": "角色活动祈愿-2", # 通常和301合并统计
            "302": "武器活动祈愿",
            "200": "常驻祈愿",
            "100": "新手祈愿"
        }
        # 用于存储统计结果
        self.stats = {
            "character_event": [], # 301 + 400
            "weapon_event": [],    # 302
            "standard": [],        # 200
            "beginners": []        # 100
        }

    def load_data(self):
        """加载JSON数据"""
        if not os.path.exists(self.file_path):
            print(f"错误: 找不到文件 {self.file_path}")
            return False
        
        try:
            with open(self.file_path, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
                self.wishes = self.data.get("list", [])
                self.uid = self.data.get("info", {}).get("uid", "未知UID")
                # 修复: 某些工具导出的info.export_time可能不准，取list中最新的时间
                if self.wishes:
                    self.latest_time = self.wishes[0].get("time")
                else:
                    self.latest_time = self.data.get("info", {}).get("export_time")
            return True
        except Exception as e:
            print(f"读取JSON失败: {e}")
            return False

    def process_wishes(self):
        """处理和分类抽卡记录"""
        # 按时间正序排列（从旧到新），方便计算保底
        sorted_wishes = sorted(self.wishes, key=lambda x: x['id']) 

        for wish in sorted_wishes:
            g_type = wish.get("uigf_gacha_type", "301") # 使用 uigf 标准字段
            
            # 将记录归类
            target_list = None
            if g_type in ["301", "400"]:
                target_list = self.stats["character_event"]
            elif g_type == "302":
                target_list = self.stats["weapon_event"]
            elif g_type == "200":
                target_list = self.stats["standard"]
            else:
                target_list = self.stats["beginners"]
            
            target_list.append(wish)

    def analyze_pool(self, pool_wishes, pool_name):
        """分析单个卡池的数据"""
        total = len(pool_wishes)
        if total == 0:
            return f"## {pool_name}\n\n无抽卡记录。\n"

        five_stars = []
        four_stars = 0
        pity_counter = 0
        
        # 遍历计算保底
        processed_wishes = [] # 存储带保底信息的记录
        
        for w in pool_wishes:
            pity_counter += 1
            rank = w.get("rank_type")
            name = w.get("name")
            
            w_info = {
                "name": name,
                "rank": rank,
                "time": w.get("time"),
                "pity": pity_counter,
                "type": w.get("item_type")
            }
            
            if rank == "5":
                five_stars.append(w_info)
                pity_counter = 0 # 重置保底
            elif rank == "4":
                four_stars += 1
            
            processed_wishes.append(w_info)

        # 倒序，为了报告中显示最新的在前面
        five_stars.reverse()
        current_pity = pity_counter
        avg_pity = round(sum(item['pity'] for item in five_stars) / len(five_stars), 2) if five_stars else 0

        # 生成 Markdown 内容
        md = f"## {pool_name}\n\n"
        md += f"- **总抽数**: `{total}`\n"
        md += f"- **五星数量**: `{len(five_stars)}` (平均出金: `{avg_pity}` 抽)\n"
        md += f"- **四星数量**: `{four_stars}`\n"
        md += f"- **当前水位**: **{current_pity}**\n\n"

        if five_stars:
            md += "### 🌟 五星获取记录\n\n"
            md += "| 角色/武器 | 类型 | 保底花费 | 获取时间 |\n"
            md += "| :--- | :---: | :---: | :--- |\n"
            for item in five_stars:
                icon = "🗡️" if item['type'] == "武器" else "👤"
                # 高亮欧皇时刻 (小于40抽)
                pity_display = f"**{item['pity']}**" if item['pity'] < 40 else f"{item['pity']}"
                if item['pity'] > 75: pity_display = f"`{item['pity']}`" # 吃保底灰色显示
                
                md += f"| {icon} {item['name']} | {item['type']} | {pity_display} | {item['time']} |\n"
        else:
            md += "> 暂无五星记录\n"
        
        md += "\n"
        return md

    def generate_report(self, output_file="genshin_wish_report.md"):
        """生成最终 Markdown 报告"""
        if not self.data:
            return

        content = f"# ✨ 原神祈愿记录分析报告\n\n"
        content += f"> **UID**: {self.uid}  \n"
        content += f"> **统计时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  \n"
        content += f"> **数据最后更新**: {self.latest_time}\n\n"
        
        content += "---\n\n"

        # 1. 角色活动祈愿分析
        content += self.analyze_pool(self.stats["character_event"], "🎭 角色活动祈愿")
        
        # 2. 武器活动祈愿分析
        content += self.analyze_pool(self.stats["weapon_event"], "⚔️ 武器活动祈愿")
        
        # 3. 常驻祈愿分析
        content += self.analyze_pool(self.stats["standard"], "🌌 常驻祈愿")

        # 4. 账号练度概览 (基于 characters 字段)
        chars = self.data.get("characters", {})
        if chars:
            content += "## 📊 账号角色概览\n\n"
            # 找出命座最高的几个角色
            sorted_chars = sorted(chars.items(), key=lambda x: x[1].get('wish', 0), reverse=True)
            
            content += "| 角色 | 抽取次数 (命座参考) | 角色 | 抽取次数 (命座参考) |\n"
            content += "| :--- | :---: | :--- | :---: |\n"
            
            # 双列显示
            for i in range(0, len(sorted_chars), 2):
                c1 = sorted_chars[i]
                row = f"| {c1[0]} | {c1[1].get('wish', 0)} "
                
                if i + 1 < len(sorted_chars):
                    c2 = sorted_chars[i+1]
                    row += f"| {c2[0]} | {c2[1].get('wish', 0)} |\n"
                else:
                    row += "| - | - |\n"
                content += row
        
        # 写入文件
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"✅ 成功! 报告已生成: {os.path.abspath(output_file)}")
        except Exception as e:
            print(f"写入文件失败: {e}")

# 使用示例
if __name__ == "__main__":
    # 这里假设文件名是 'data.json'，你需要改成你实际的文件名
    # 或者直接把 JSON 内容粘贴到一个名为 paimon_moe_local_data.json 的文件中
    json_file = r"C:\Users\ricar\Downloads\paimon-moe-local-data (1).json"
    
    # 为了方便演示，如果没有文件，我会提示创建
    if not os.path.exists(json_file):
        print(f"请确保目录中存在 {json_file} 文件，或者修改代码中的文件名。")
    else:
        analyzer = GenshinWishAnalyzer(json_file)
        if analyzer.load_data():
            analyzer.process_wishes()
            analyzer.generate_report()