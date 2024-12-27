import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import os

class ACLGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("ACL命令生成器")
        
        # 创建Excel文件（如果不存在）
        self.excel_file = "acl_records.xlsx"
        self.create_excel_if_not_exists()
        
        # 创建界面元素
        self.create_widgets()
        
    def create_excel_if_not_exists(self):
        if not os.path.exists(self.excel_file):
            # 创建带有默认楼层工作表的Excel文件
            floors = ['1F']  # 默认只创建一个楼层
            with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                for floor in floors:
                    df = pd.DataFrame(columns=['时间', '源IP', '目标IP', '端口', 'ACL编号', 'ACL命令'])
                    df.to_excel(writer, sheet_name=floor, index=False)
        else:
            # 检查文件是否可以正常打开
            try:
                pd.read_excel(self.excel_file)
            except Exception as e:
                # 如果文件损坏或无法打开，创建新文件
                floors = ['1F']
                with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                    for floor in floors:
                        df = pd.DataFrame(columns=['时间', '源IP', '目标IP', '端口', 'ACL编号', 'ACL命令'])
                        df.to_excel(writer, sheet_name=floor, index=False)
    
    def create_widgets(self):
        # 创建左右分栏
        left_frame = ttk.Frame(self.root)
        left_frame.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')
        
        right_frame = ttk.Frame(self.root)
        right_frame.grid(row=0, column=1, padx=10, pady=5, sticky='nsew')
        
        # 关于按钮（放在右上角）
        self.about_btn = ttk.Button(self.root, text="关于", command=self.show_about)
        self.about_btn.grid(row=0, column=2, padx=5, pady=5, sticky='ne')
        
        # 左侧：楼层管理
        ttk.Label(left_frame, text="楼层管理").grid(row=0, column=0, columnspan=2, pady=5)
        
        # 楼层列表
        self.floor_listbox = tk.Listbox(left_frame, height=10, width=20)
        self.floor_listbox.grid(row=1, column=0, columnspan=2, pady=5)
        
        # 添加楼层输入框和按钮
        self.new_floor = ttk.Entry(left_frame, width=10)
        self.new_floor.grid(row=2, column=0, pady=5)
        ttk.Button(left_frame, text="添加楼层", command=self.add_floor).grid(row=2, column=1, pady=5)
        
        # 删除楼层按钮
        ttk.Button(left_frame, text="删除楼层", command=self.delete_floor).grid(row=3, column=0, columnspan=2, pady=5)
        
        # 右侧：ACL生成
        ttk.Label(right_frame, text="ACL命令生成").grid(row=0, column=0, columnspan=2, pady=5)
        
        # 楼层选择（使用当前楼层列表）
        ttk.Label(right_frame, text="选择楼层:").grid(row=1, column=0, padx=5, pady=5)
        self.floor_var = tk.StringVar()
        self.floor_combo = ttk.Combobox(right_frame, textvariable=self.floor_var)
        self.floor_combo.grid(row=1, column=1, padx=5, pady=5)
        
        # 源IP输入
        ttk.Label(right_frame, text="源IP:").grid(row=2, column=0, padx=5, pady=5)
        self.source_ip = ttk.Entry(right_frame)
        self.source_ip.grid(row=2, column=1, padx=5, pady=5)
        
        # 目标IP输入
        ttk.Label(right_frame, text="目标IP:").grid(row=3, column=0, padx=5, pady=5)
        self.dest_ip = ttk.Entry(right_frame)
        self.dest_ip.grid(row=3, column=1, padx=5, pady=5)
        
        # 端口号输入
        ttk.Label(right_frame, text="端口号:").grid(row=4, column=0, padx=5, pady=5)
        self.port = ttk.Entry(right_frame)
        self.port.insert(0, "0")
        self.port.grid(row=4, column=1, padx=5, pady=5)
        
        # 生成按钮（恢复原位置）
        self.generate_btn = ttk.Button(right_frame, text="生成ACL命令", command=self.generate_acl)
        self.generate_btn.grid(row=5, column=0, columnspan=2, pady=10)
        
        # 结果显示区域
        self.result_text = tk.Text(right_frame, height=5, width=50)
        self.result_text.grid(row=6, column=0, columnspan=2, padx=5, pady=5)
        
        # 最后更新楼层列表
        self.update_floor_list()

    def update_floor_list(self):
        """更新楼层列表和下拉框"""
        try:
            with pd.ExcelFile(self.excel_file) as xls:
                floors = xls.sheet_names
                
            # 更新列表框
            self.floor_listbox.delete(0, tk.END)
            for floor in floors:
                self.floor_listbox.insert(tk.END, floor)
                
            # 更新下拉框
            self.floor_combo['values'] = floors
            if floors:
                self.floor_combo.set(floors[0])
        except Exception as e:
            messagebox.showerror("错误", f"更新楼层列表失败：{str(e)}")

    def add_floor(self):
        """添加新楼层"""
        new_floor = self.new_floor.get().strip()
        if not new_floor:
            messagebox.showerror("错误", "请输入楼层名称！")
            return
            
        try:
            # 检查楼层是否已存在
            with pd.ExcelFile(self.excel_file) as xls:
                if new_floor in xls.sheet_names:
                    messagebox.showerror("错误", "该楼层已存在！")
                    return
            
            # 创建新的工作表
            df = pd.DataFrame(columns=['时间', '源IP', '目标IP', '端口', 'ACL编号', 'ACL命令'])
            with pd.ExcelWriter(self.excel_file, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=new_floor, index=False)
            
            # 清空输入框并更新列表
            self.new_floor.delete(0, tk.END)
            self.update_floor_list()
            messagebox.showinfo("成功", f"已添加楼层：{new_floor}")
        except Exception as e:
            messagebox.showerror("错误", f"添加楼层失败：{str(e)}")

    def delete_floor(self):
        """删除选中的楼层"""
        selection = self.floor_listbox.curselection()
        if not selection:
            messagebox.showerror("错误", "请先选择要删除的楼层！")
            return
            
        floor = self.floor_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除楼层 {floor} 吗？"):
            try:
                # 读取所有工作表
                with pd.ExcelFile(self.excel_file) as xls:
                    sheets = xls.sheet_names
                    if len(sheets) <= 1:
                        messagebox.showerror("错误", "不能删除最后一个楼层！")
                        return
                        
                    # 读取除了要删除的楼层外的所有数据
                    dfs = {sheet: pd.read_excel(self.excel_file, sheet_name=sheet)
                          for sheet in sheets if sheet != floor}
                
                # 重新写入Excel文件
                with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                    for sheet_name, df in dfs.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                self.update_floor_list()
                messagebox.showinfo("成功", f"已删除楼层：{floor}")
            except Exception as e:
                messagebox.showerror("错误", f"删除楼层失败：{str(e)}")

    def get_next_acl_number(self, floor, target_ip):
        try:
            df = pd.read_excel(self.excel_file, sheet_name=floor)
            if df.empty:
                return 1
            
            # 筛选出目标IP相关的记录
            target_records = df[df['目标IP'] == target_ip]
            if target_records.empty:
                # 如果没有该目标IP的记录，获取所有记录中最大编号+1
                existing_numbers = df['ACL编号'].dropna().astype(int)
                if len(existing_numbers) == 0:
                    return 1
                return max(existing_numbers) + 1
            else:
                # 如果有该目标IP的记录，获取该IP相关记录的最大编号+1
                existing_numbers = target_records['ACL编号'].dropna().astype(int)
                return max(existing_numbers) + 1
        except:
            return 1
            
    def check_existing_acl(self, floor, source, dest):
        """检查是否存在相同源IP和目标IP的ACL"""
        try:
            df = pd.read_excel(self.excel_file, sheet_name=floor)
            if df.empty:
                return False, None
            
            # 检查是否存在相同的源IP和目标IP组合
            existing = df[
                ((df['源IP'] == source) & (df['目标IP'] == dest)) |
                ((df['源IP'] == dest) & (df['目标IP'] == source))
            ]
            
            if not existing.empty:
                # 如果找到匹配的记录，返回True和第一条匹配记录的ACL编号
                return True, existing.iloc[0]['ACL编号']
            return False, None
        except:
            return False, None

    def generate_acl(self):
        # 获取输入值
        floor = self.floor_var.get()
        source = self.source_ip.get()
        dest = self.dest_ip.get()
        port_num = self.port.get() or "0"
        
        # 基本验证
        if not all([floor, source, dest]):
            messagebox.showerror("错误", "请填写楼层、源IP和目标IP！")
            return
            
        # 检查是否存在相同的ACL
        exists, existing_number = self.check_existing_acl(floor, source, dest)
        if exists:
            messagebox.showwarning(
                "警告", 
                f"该源IP和目标IP的ACL规则已存在！\n规则编号: {existing_number}"
            )
            return
            
        # 只获取一个ACL编号（源到目标的规则编号）
        acl_number = self.get_next_acl_number(floor, dest)
        
        # 生成双向的ACL命令，使用相同的rule编号
        if port_num == "0":
            acl_command1 = f"rule {acl_number} permit ip source {source} 0 destination {dest} 0"
            acl_command2 = f"rule {acl_number} permit ip source {dest} 0 destination {source} 0"
        else:
            acl_command1 = f"rule {acl_number} permit ip source {source} 0 destination {dest} 0 destination-port eq {port_num}"
            acl_command2 = f"rule {acl_number} permit ip source {dest} 0 destination {source} 0 destination-port eq {port_num}"
        
        # 合并两条命令
        combined_commands = f"{acl_command1}\n{acl_command2}"
        
        # 显示在界面上
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, combined_commands)
        
        # 保存到Excel
        try:
            df = pd.read_excel(self.excel_file, sheet_name=floor)
            new_rows = pd.DataFrame([
                {
                    '时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    '源IP': source,
                    '目标IP': dest,
                    '端口': port_num,
                    'ACL编号': acl_number,
                    'ACL命令': acl_command1
                },
                {
                    '时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    '源IP': dest,
                    '目标IP': source,
                    '端口': port_num,
                    'ACL编号': acl_number,
                    'ACL命令': acl_command2
                }
            ])
            
            # 合并数据框并按ACL编号排序
            df = pd.concat([df, new_rows], ignore_index=True)
            df = df.sort_values(by=['ACL编号', '源IP'])
            df = df.reset_index(drop=True)
            
            # 使用 ExcelWriter 保存所有工作表
            with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='a') as writer:
                # 删除已存在的工作表
                if floor in writer.book.sheetnames:
                    idx = writer.book.sheetnames.index(floor)
                    writer.book.remove(writer.book.worksheets[idx])
                
                # 写入排序后的工作表
                df.to_excel(writer, sheet_name=floor, index=False)
                
                # 自动调整列宽
                worksheet = writer.sheets[floor]
                for idx, col in enumerate(df.columns):
                    # 获取列中最长的内容长度
                    max_length = max(
                        df[col].astype(str).apply(len).max(),  # 数据的最大长度
                        len(str(col))  # 列名的长度
                    )
                    # 设置列宽（稍微加宽一点，使显示更美观）
                    adjusted_width = (max_length + 2) * 1.2
                    # 设置列宽
                    worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width
                
            messagebox.showinfo("成功", "ACL命令已生成并保存到Excel文件中！")
        except Exception as e:
            messagebox.showerror("错误", f"保存到Excel时出错：{str(e)}")
            print(f"详细错误信息: {str(e)}")

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo("关于", "作者：细雨\nQQ：2082216455")

if __name__ == "__main__":
    root = tk.Tk()
    app = ACLGenerator(root)
    root.mainloop() 