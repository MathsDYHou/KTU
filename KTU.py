import numpy as np
import pandas as pd
from sklearn.preprocessing import StandardScaler
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from threading import Thread
import sys
from ctypes import windll
class KMOAnalyzerApp:
    def __init__(self,root):
        self.root=root
        self.root.title("纤维植被毯修复效果评价系统V1.0")
        self.root.geometry("1000x700")
        if sys.platform=='win32':
            try:
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass
        main_container=tk.Frame(self.root)
        main_container.pack(fill="both",expand=True)
        canvas_container=tk.Frame(main_container)
        canvas_container.pack(side="top",fill="both",expand=True)
        self.canvas=tk.Canvas(canvas_container,highlightthickness=0,width=1000)
        scrollbar=ttk.Scrollbar(canvas_container,orient="vertical",command=self.canvas.yview)
        self.scroll_frame=tk.Frame(self.canvas)
        self.scroll_frame.bind("<Configure>",lambda e:self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0),window=self.scroll_frame,anchor="nw",tags="frame")
        self.canvas.bind("<Configure>",self._on_canvas_resize)
        scrollbar.pack(side="right",fill="y")
        self.canvas.pack(side="left",fill="both",expand=True)
        self.scroll_frame.bind("<Enter>",lambda e:self._bind_mousewheel())
        self.scroll_frame.bind("<Leave>",lambda e:self._unbind_mousewheel())
        self.title_font=('微软雅黑',14,'bold')
        self.label_font=('微软雅黑',12)
        self.btn_font=('微软雅黑',12,'bold')
        self.listbox_font=('微软雅黑',12)
        self.file_path=tk.StringVar()
        self.kmo_threshold=tk.DoubleVar(value=0.6)
        self.results={}
        self.create_widgets()
    def _on_canvas_resize(self,event):
        canvas_width=event.width
        self.canvas.itemconfigure("frame",width=canvas_width)
    def _bind_mousewheel(self):
        self.canvas.bind_all("<MouseWheel>",lambda e:self.canvas.yview_scroll(-int(e.delta/120),"units"))
    def _unbind_mousewheel(self):
        self.canvas.unbind_all("<MouseWheel>")
    def create_widgets(self):
        style=ttk.Style()
        style.configure('Title.TLabelframe.Label',font=self.title_font)
        style.configure('Large.TButton',font=self.btn_font,padding=10)
        file_frame=ttk.LabelFrame(self.scroll_frame,text="数据文件选择",style='Title.TLabelframe',padding=20)
        file_frame.pack(pady=20,padx=20,fill=tk.X)
        ttk.Label(file_frame,text="Excel文件路径:",font=self.label_font).grid(row=0,column=0,padx=20,sticky='e')
        ttk.Entry(file_frame,textvariable=self.file_path,width=50,font=self.listbox_font).grid(row=0,column=1,padx=20,sticky='ew')
        ttk.Button(file_frame,text="浏览",command=self.browse_file,style='Large.TButton').grid(row=0,column=2,padx=20)
        param_frame=ttk.LabelFrame(self.scroll_frame,text="分析参数设置",style='Title.TLabelframe',padding=20)
        param_frame.pack(pady=10,padx=20,fill=tk.X)
        kmo_frame=ttk.Frame(param_frame)
        kmo_frame.pack(side=tk.LEFT,padx=20)
        ttk.Label(kmo_frame,text="KMO阈值:",font=self.label_font).pack(side=tk.LEFT)
        ttk.Entry(kmo_frame,textvariable=self.kmo_threshold,width=10,font=self.label_font).pack(side=tk.LEFT,padx=10)
        negative_frame=ttk.Frame(param_frame)
        negative_frame.pack(side=tk.RIGHT,padx=20)
        ttk.Label(negative_frame,text="逆向化指标选取",font=('微软雅黑',12),anchor="center").pack(pady=5,fill=tk.X)
        self.negative_listbox=tk.Listbox(negative_frame,selectmode=tk.MULTIPLE,height=10,width=40,font=self.listbox_font,relief='groove',selectbackground='#0078D7',selectforeground='white',activestyle='dotbox')
        self.negative_listbox.pack(pady=10)
        btn_group=ttk.Frame(negative_frame)
        btn_group.pack(pady=5)
        ttk.Button(btn_group,text="全选",command=self.select_all,style='Large.TButton').pack(side=tk.LEFT,padx=5)
        ttk.Button(btn_group,text="重置",command=self.reset_selection,style='Large.TButton').pack(side=tk.LEFT,padx=5)
        ttk.Button(btn_group,text="反选",command=self.invert_selection,style='Large.TButton').pack(side=tk.LEFT,padx=5)
        result_frame=ttk.LabelFrame(self.scroll_frame,text="分析结果",style='Title.TLabelframe',padding=20)
        result_frame.pack(pady=20,padx=20,fill=tk.BOTH,expand=True)
        result_text_scroll=ttk.Scrollbar(result_frame)
        result_text_scroll.pack(side=tk.RIGHT,fill=tk.Y)
        self.result_text=tk.Text(result_frame,wrap=tk.WORD,yscrollcommand=result_text_scroll.set,font=self.listbox_font,padx=20,pady=20)
        self.result_text.pack(fill=tk.BOTH,expand=True)
        result_text_scroll.config(command=self.result_text.yview)
        btn_frame=ttk.Frame(self.scroll_frame)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame,text="开始分析",command=self.start_analysis,style='Large.TButton').pack(side=tk.LEFT,padx=20)
        ttk.Button(btn_frame,text="导出结果",command=self.export_results,style='Large.TButton').pack(side=tk.LEFT,padx=20)
        ttk.Button(btn_frame,text="退出系统",command=self.root.quit,style='Large.TButton').pack(side=tk.LEFT,padx=20)
    def browse_file(self):
        file_path=filedialog.askopenfilename(filetypes=[("Excel文件","*.xlsx")])
        if file_path:
            self.file_path.set(file_path)
            self.load_columns_preview()
    def load_columns_preview(self):
        try:
            df=pd.read_excel(self.file_path.get(),nrows=1)
            self.negative_listbox.delete(0,tk.END)
            for col in df.select_dtypes(include=np.number).columns:
                self.negative_listbox.insert(tk.END,col)
        except Exception as e:
            messagebox.showerror("错误",f"读取文件失败:{str(e)}")
    def select_all(self):
        self.negative_listbox.selection_set(0,tk.END)
    def reset_selection(self):
        self.negative_listbox.selection_clear(0,tk.END)
    def invert_selection(self):
        all_items=range(self.negative_listbox.size())
        current_selection=set(self.negative_listbox.curselection())
        for i in all_items:
            if i in current_selection:
                self.negative_listbox.selection_clear(i)
            else:
                self.negative_listbox.selection_set(i)
    def start_analysis(self):
        if not self.file_path.get():
            messagebox.showwarning("警告","请先选择Excel文件")
            return
        Thread(target=self.run_analysis,daemon=True).start()
    def run_analysis(self):
        try:
            self.result_text.delete(1.0,tk.END)
            self.result_text.insert(tk.END,"正在读取数据...\n")
            self.root.update()
            data=pd.read_excel(self.file_path.get(),index_col=0)
            data['实验分组']=data.index
            data.reset_index(drop=True,inplace=True)
            selected_indices=self.negative_listbox.curselection()
            negative_columns=[self.negative_listbox.get(i) for i in selected_indices]
            numeric_cols=data.select_dtypes(include=[np.number]).columns.tolist()
            data_numeric=data[numeric_cols]
            scaler=StandardScaler()
            data_standardized=pd.DataFrame(scaler.fit_transform(data_numeric),columns=data_numeric.columns)
            self.result_text.insert(tk.END,"\n正在进行KMO指标筛选...\n")
            final_data,final_kmo=self.iterative_kmo_screening(data_standardized,self.kmo_threshold.get())
            self.result_text.insert(tk.END,"\n正在计算指标权重...\n")
            weight=self.entropy_weight(final_data,negative_columns)
            self.result_text.insert(tk.END,"\n正在进行TOPSIS评分...\n")
            score=self.topsis(final_data,weight,negative_columns)
            data['TOPSIS评分']=score
            data['总体排名']=data['TOPSIS评分'].rank(ascending=False).astype(int)
            data['实验分组排名']=data.groupby('实验分组')['TOPSIS评分'].rank(ascending=False).astype(int)
            self.results={'data':data,'kmo':final_kmo,'weights':dict(zip(final_data.columns,weight)),'selected_columns':final_data.columns.tolist()}
            self.show_results()
            messagebox.showinfo("完成","分析完成！")
        except Exception as e:
            messagebox.showerror("错误",f"分析错误:{str(e)}")
            self.result_text.insert(tk.END,f"\n错误:{str(e)}")
    def show_results(self):
        self.result_text.delete(1.0,tk.END)
        self.result_text.insert(tk.END,"====分析结果====\n")
        self.result_text.insert(tk.END,f"最终KMO值:{self.results['kmo']:.4f}\n")
        self.result_text.insert(tk.END,f"保留指标({len(self.results['selected_columns'])}个):\n")
        self.result_text.insert(tk.END,"\n".join(self.results['selected_columns'])+"\n\n")
        self.result_text.insert(tk.END,"====指标权重====\n")
        for col,w in self.results['weights'].items():
            self.result_text.insert(tk.END,f"{col}:{w:.4f}\n")
        data=self.results['data']
        max_indices=data.groupby('实验分组')['TOPSIS评分'].idxmax()
        top_groups=data.loc[max_indices]
        self.result_text.insert(tk.END,"\n====各实验分组最高得分组别====\n")
        for _,row in top_groups.iterrows():
            self.result_text.insert(tk.END,f"实验分组：{row['实验分组']} | 最高组别：{row['组别']} | 得分：{row['TOPSIS评分']:.4f}\n")
        self.result_text.insert(tk.END,"\n====综合排名====\n")
        output_data=self.results['data'][['实验分组','组别','TOPSIS评分','总体排名','实验分组排名']]
        for _,row in output_data.sort_values('总体排名').iterrows():
            self.result_text.insert(tk.END,f"{row['实验分组']} | {row['组别']} | 得分:{row['TOPSIS评分']:.4f} | "f"总排名:{row['总体排名']} | 分组排名:{row['实验分组排名']}\n")
    def export_results(self):
        if not self.results:
            messagebox.showwarning("警告","请先完成分析")
            return
        save_path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel文件","*.xlsx")])
        if save_path:
            try:
                with pd.ExcelWriter(save_path) as writer:
                    self.results['data'].to_excel(writer,sheet_name='综合结果')
                    pd.DataFrame.from_dict(self.results['weights'],orient='index',columns=['权重']).to_excel(writer,sheet_name='指标权重')
                messagebox.showinfo("成功",f"结果已保存到:{save_path}")
            except Exception as e:
                messagebox.showerror("错误",f"保存失败:{str(e)}")
    def calculate_kmo(self,data):
        corr_matrix=np.corrcoef(data,rowvar=False)
        inv_corr=np.linalg.inv(corr_matrix)
        pcorr_matrix=-inv_corr/np.sqrt(np.outer(np.diag(inv_corr),np.diag(inv_corr)))
        sum_sq_corr=np.sum(corr_matrix**2)-np.sum(np.diag(corr_matrix**2))
        sum_sq_pcorr=np.sum(pcorr_matrix**2)-np.sum(np.diag(pcorr_matrix**2))
        return sum_sq_corr/(sum_sq_corr+sum_sq_pcorr)
    def iterative_kmo_screening(self,data,threshold):
        current_data=data.copy()
        columns=current_data.columns.tolist()
        while True:
            kmo=self.calculate_kmo(current_data.values)
            if kmo>=threshold:
                break
            corr_matrix=np.corrcoef(current_data,rowvar=False)
            np.fill_diagonal(corr_matrix,0)
            max_idx=np.unravel_index(np.argmax(np.abs(corr_matrix)),corr_matrix.shape)
            var1,var2=max_idx[0],max_idx[1]
            avg_corr_var1=np.mean(np.abs(corr_matrix[var1,:]))
            avg_corr_var2=np.mean(np.abs(corr_matrix[var2,:]))
            drop_col=columns[var1] if avg_corr_var1>=avg_corr_var2 else columns[var2]
            columns.remove(drop_col)
            current_data=data[columns]
        return current_data,kmo
    def entropy_weight(self,data,negative_cols):
        data_processed=data.copy()
        for col in negative_cols:
            if col in data_processed.columns:
                data_processed[col]=1/(data_processed[col]+1e-6)
        data_norm=(data_processed-np.min(data_processed,axis=0))/(np.max(data_processed,axis=0)-np.min(data_processed,axis=0)+1e-6)*0.999+0.001
        p=data_norm/np.sum(data_norm,axis=0)
        entropy=-np.sum(p*np.log(p+1e-10),axis=0)/np.log(len(data))
        return (1-entropy)/np.sum(1-entropy)
    def topsis(self,data,weight,negative_cols):
        columns=data.columns.tolist()
        negative_indices=[columns.index(col) for col in negative_cols if col in columns]
        data_processed=data.values.copy().astype(float)
        for idx in negative_indices:
            data_processed[:,idx]=1/(data_processed[:,idx]+1e-6)
        weight=np.array(weight).reshape(1,-1)
        weighted_data=data_processed*weight
        ideal_best,ideal_worst=[],[]
        for i in range(weighted_data.shape[1]):
            if i in negative_indices:
                ideal_best.append(np.min(weighted_data[:,i]))
                ideal_worst.append(np.max(weighted_data[:,i]))
            else:
                ideal_best.append(np.max(weighted_data[:,i]))
                ideal_worst.append(np.min(weighted_data[:,i]))
        dist_best=np.sqrt(np.sum((weighted_data-ideal_best)**2,axis=1))
        dist_worst=np.sqrt(np.sum((weighted_data-ideal_worst)**2,axis=1))
        return dist_worst/(dist_best+dist_worst+1e-10)
if __name__=="__main__":
    root=tk.Tk()
    app=KMOAnalyzerApp(root)
    root.mainloop()