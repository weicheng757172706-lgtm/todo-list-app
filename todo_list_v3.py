#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Todo List Application - Stable Version
Features: Excel-like table, priority matrix, no external dependencies
"""

import json
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox


class TodoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Todo List - 每日任务管理")
        self.root.geometry("1100x750")
        
        # Data file path - save in same directory as exe
        if getattr(sys, 'frozen', False):
            # Running as exe
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        self.data_file = os.path.join(base_path, "todo_data.json")
        
        # Priority options (Eisenhower Matrix)
        self.priority_options = ["重要紧急", "重要不紧急", "不重要紧急", "不重要不紧急"]
        self.priority_colors = {
            "重要紧急": "#FFE5E5",
            "重要不紧急": "#E5F9F6",
            "不重要紧急": "#FFF9E5",
            "不重要不紧急": "#E5F9E5"
        }
        
        # Load data
        self.data = self.load_data()
        
        # Setup GUI
        self.setup_gui()
        
        # Refresh display
        self.refresh_display()
    
    def load_data(self):
        """Load data from JSON file"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading data: {e}")
                pass
        return {"tasks": [], "completed": []}
    
    def save_data(self):
        """Save data to JSON file"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving data: {e}")
            messagebox.showerror("错误", f"保存数据失败:\n{str(e)}")
    
    def setup_gui(self):
        """Setup the GUI components"""
        # Title
        title_frame = tk.Frame(self.root, bg="#4F81BD", height=60)
        title_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E))
        title_frame.grid_propagate(False)
        
        title_label = tk.Label(title_frame, text="📝 Todo List - 每日任务管理", 
                               font=("Microsoft YaHei", 16, "bold"),
                               bg="#4F81BD", fg="white")
        title_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Input frame
        input_frame = tk.LabelFrame(self.root, text="添加新任务", font=("Microsoft YaHei", 10),
                                    padx=10, pady=10)
        input_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=10, pady=5)
        
        tk.Label(input_frame, text="任务内容:", font=("Microsoft YaHei", 9)).grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        self.task_entry = tk.Entry(input_frame, width=45, font=("Microsoft YaHei", 10))
        self.task_entry.grid(row=0, column=1, padx=(0, 10))
        self.task_entry.bind('<Return>', lambda e: self.add_task())
        
        tk.Label(input_frame, text="重要紧急程度:", font=("Microsoft YaHei", 9)).grid(row=0, column=2, padx=(10, 5), sticky=tk.W)
        self.priority_var = tk.StringVar(value=self.priority_options[0])
        self.priority_combo = ttk.Combobox(input_frame, textvariable=self.priority_var,
                                            values=self.priority_options, state='readonly',
                                            width=18, font=("Microsoft YaHei", 9))
        self.priority_combo.grid(row=0, column=3, padx=(0, 10))
        
        add_btn = tk.Button(input_frame, text="✚ 添加", command=self.add_task,
                             bg="#4F81BD", fg="white", font=("Microsoft YaHei", 9, "bold"),
                             cursor="hand2", padx=20)
        add_btn.grid(row=0, column=4, padx=(10, 0))
        
        # Notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), 
                      padx=10, pady=10)
        
        # Pending tasks tab
        pending_frame = tk.Frame(notebook, padx=10, pady=10)
        notebook.add(pending_frame, text="📋 待完成任务")
        self.setup_pending_tab(pending_frame)
        
        # Completed tasks tab
        completed_frame = tk.Frame(notebook, padx=10, pady=10)
        notebook.add(completed_frame, text="✅ 已完成任务")
        self.setup_completed_tab(completed_frame)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)
    
    def setup_pending_tab(self, parent):
        """Setup pending tasks tab"""
        # Toolbar
        toolbar = tk.Frame(parent)
        toolbar.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        complete_btn = tk.Button(toolbar, text="✓ 标记完成", command=self.complete_task,
                                 bg="#5CB85C", fg="white", font=("Microsoft YaHei", 9),
                                 cursor="hand2", padx=15)
        complete_btn.grid(row=0, column=0, padx=(0, 5))
        
        delete_btn = tk.Button(toolbar, text="🗑 删除", command=self.delete_task,
                               bg="#D9534F", fg="white", font=("Microsoft YaHei", 9),
                               cursor="hand2", padx=15)
        delete_btn.grid(row=0, column=1, padx=5)
        
        refresh_btn = tk.Button(toolbar, text="🔄 刷新", command=self.refresh_display,
                                bg="#5BC0DE", fg="white", font=("Microsoft YaHei", 9),
                                cursor="hand2", padx=15)
        refresh_btn.grid(row=0, column=2, padx=5)
        
        # Table
        table_frame = tk.Frame(parent)
        table_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        columns = ('ID', '任务内容', '重要紧急程度', '创建时间')
        self.pending_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        col_widths = {'ID': 60, '任务内容': 500, '重要紧急程度': 130, '创建时间': 160}
        for col in columns:
            self.pending_tree.heading(col, text=col, anchor=tk.CENTER)
            self.pending_tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)
        
        self.pending_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.pending_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.pending_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.pending_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.pending_tree.configure(xscrollcommand=h_scrollbar.set)
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
    
    def setup_completed_tab(self, parent):
        """Setup completed tasks tab"""
        # Filter frame
        filter_frame = tk.Frame(parent)
        filter_frame.grid(row=0, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        tk.Label(filter_frame, text="筛选日期:", font=("Microsoft YaHei", 9)).grid(row=0, column=0, padx=(0, 10))
        
        self.date_filter = ttk.Combobox(filter_frame, width=20, state='readonly', font=("Microsoft YaHei", 9))
        self.date_filter.grid(row=0, column=1, padx=(0, 10))
        self.date_filter.bind('<<ComboboxSelected>>', lambda e: self.filter_completed())
        
        # Toolbar
        toolbar = tk.Frame(parent)
        toolbar.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 10))
        
        delete_completed_btn = tk.Button(toolbar, text="🗑 删除记录", command=self.delete_completed,
                                         bg="#D9534F", fg="white", font=("Microsoft YaHei", 9),
                                         cursor="hand2", padx=15)
        delete_completed_btn.grid(row=0, column=0, padx=(0, 5))
        
        refresh_completed_btn = tk.Button(toolbar, text="🔄 刷新", command=self.refresh_completed,
                                          bg="#5BC0DE", fg="white", font=("Microsoft YaHei", 9),
                                          cursor="hand2", padx=15)
        refresh_completed_btn.grid(row=0, column=1, padx=5)
        
        # Table
        table_frame = tk.Frame(parent)
        table_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        columns = ('ID', '任务内容', '重要紧急程度', '完成时间')
        self.completed_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)
        
        col_widths = {'ID': 60, '任务内容': 500, '重要紧急程度': 130, '完成时间': 160}
        for col in columns:
            self.completed_tree.heading(col, text=col, anchor=tk.CENTER)
            self.completed_tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)
        
        self.completed_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.completed_tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.completed_tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.completed_tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.completed_tree.configure(xscrollcommand=h_scrollbar.set)
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
    
    def add_task(self):
        """Add a new task"""
        task_text = self.task_entry.get().strip()
        if not task_text:
            messagebox.showwarning("警告", "请输入任务内容！")
            return
        
        priority = self.priority_var.get()
        
        task = {
            "id": len(self.data["tasks"]) + len(self.data["completed"]) + 1,
            "text": task_text,
            "priority": priority,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        self.data["tasks"].append(task)
        self.save_data()
        self.task_entry.delete(0, tk.END)
        self.refresh_display()
        messagebox.showinfo("成功", "任务已添加！")
    
    def complete_task(self):
        """Mark a task as completed"""
        selection = self.pending_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择要完成的任务！")
            return
        
        item = self.pending_tree.item(selection[0])
        task_id = int(item['values'][0])
        
        for i, task in enumerate(self.data["tasks"]):
            if task["id"] == task_id:
                task["completed"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.data["completed"].append(task)
                self.data["tasks"].pop(i)
                break
        
        self.save_data()
        self.refresh_display()
        messagebox.showinfo("成功", "任务已标记为完成！")
    
    def delete_task(self):
        """Delete a pending task"""
        selection = self.pending_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择要删除的任务！")
            return
        
        if messagebox.askyesno("确认", "确定要删除这个任务吗？"):
            item = self.pending_tree.item(selection[0])
            task_id = int(item['values'][0])
            
            self.data["tasks"] = [t for t in self.data["tasks"] if t["id"] != task_id]
            self.save_data()
            self.refresh_display()
            messagebox.showinfo("成功", "任务已删除！")
    
    def delete_completed(self):
        """Delete a completed task record"""
        selection = self.completed_tree.selection()
        if not selection:
            messagebox.showwarning("警告", "请选择要删除的记录！")
            return
        
        if messagebox.askyesno("确认", "确定要删除这个记录吗？"):
            item = self.completed_tree.item(selection[0])
            task_id = int(item['values'][0])
            
            self.data["completed"] = [t for t in self.data["completed"] if t["id"] != task_id]
            self.save_data()
            self.refresh_completed()
            messagebox.showinfo("成功", "记录已删除！")
    
    def refresh_display(self):
        """Refresh the display"""
        self.refresh_pending()
        self.refresh_completed()
    
    def refresh_pending(self):
        """Refresh pending tasks display"""
        self.pending_tree.delete(*self.pending_tree.get_children())
        
        for task in self.data["tasks"]:
            item = self.pending_tree.insert('', tk.END, values=(
                task["id"],
                task["text"],
                task.get("priority", ""),
                task["created"]
            ))
            
            priority = task.get("priority", "")
            if priority in self.priority_colors:
                self.pending_tree.item(item, tags=(priority,))
        
        for priority, color in self.priority_colors.items():
            self.pending_tree.tag_configure(priority, background=color)
    
    def refresh_completed(self):
        """Refresh completed tasks display"""
        self.completed_tree.delete(*self.completed_tree.get_children())
        
        dates = set()
        for task in self.data["completed"]:
            if "completed" in task:
                date = task["completed"].split()[0]
                dates.add(date)
        
        dates = sorted(list(dates), reverse=True)
        self.date_filter['values'] = ['全部'] + dates
        if not self.date_filter.get():
            self.date_filter.set('全部')
        
        self.filter_completed()
    
    def filter_completed(self):
        """Filter completed tasks by date"""
        self.completed_tree.delete(*self.completed_tree.get_children())
        
        selected_date = self.date_filter.get()
        
        for task in self.data["completed"]:
            if "completed" in task:
                task_date = task["completed"].split()[0]
                
                if selected_date == '全部' or task_date == selected_date:
                    item = self.completed_tree.insert('', tk.END, values=(
                        task["id"],
                        task["text"],
                        task.get("priority", ""),
                        task["completed"]
                    ))
                    
                    priority = task.get("priority", "")
                    if priority in self.priority_colors:
                        self.completed_tree.item(item, tags=(priority,))
        
        for priority, color in self.priority_colors.items():
            self.completed_tree.tag_configure(priority, background=color)


if __name__ == "__main__":
    import sys
    root = tk.Tk()
    app = TodoApp(root)
    root.mainloop()
