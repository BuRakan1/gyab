import tkinter as tk
from tkinter import ttk, messagebox
import datetime
from tkcalendar import DateEntry
import sqlite3
from collections import defaultdict

class AbsenceMonitoringSystem:
    def __init__(self, root, parent_app, colors, fonts):
        self.root = root
        self.parent_app = parent_app
        self.colors = colors
        self.fonts = fonts
        self.conn = parent_app.conn
        
        self.setup_window()
        self.create_widgets()
        self.load_courses()
        
    def setup_window(self):
        """إعداد نافذة نظام مراقبة الغياب"""
        self.window = tk.Toplevel(self.root)
        self.window.title("نظام مراقبة الغياب والإنذارات")
        self.window.geometry("1200x700")
        self.window.configure(bg=self.colors["light"])
        self.window.grab_set()
        
        # توسيط النافذة
        x = (self.window.winfo_screenwidth() - 1200) // 2
        y = (self.window.winfo_screenheight() - 700) // 2
        self.window.geometry(f"1200x700+{x}+{y}")
        
    def create_widgets(self):
        """إنشاء عناصر الواجهة"""
        # العنوان الرئيسي
        title_frame = tk.Frame(self.window, bg=self.colors["primary"], height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text="نظام مراقبة الغياب والإنذارات - حسب تعليمات التدريب المستديمة",
            font=self.fonts["large_title"],
            bg=self.colors["primary"],
            fg="white"
        ).pack(expand=True)
        
        # إطار اختيار الدورة
        course_frame = tk.Frame(self.window, bg=self.colors["light"], pady=10)
        course_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(
            course_frame,
            text="اختر الدورة:",
            font=self.fonts["text_bold"],
            bg=self.colors["light"]
        ).pack(side=tk.RIGHT, padx=5)
        
        self.course_var = tk.StringVar()
        self.course_combo = ttk.Combobox(
            course_frame,
            textvariable=self.course_var,
            font=self.fonts["text"],
            width=30,
            state="readonly"
        )
        self.course_combo.pack(side=tk.RIGHT, padx=5)
        self.course_combo.bind("<<ComboboxSelected>>", self.on_course_selected)
        
        # إطار معلومات الدورة المحسّن
        self.info_frame = tk.LabelFrame(
            self.window,
            text="معلومات الدورة وحدود الغياب المسموح",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=20,
            pady=20
        )
        self.info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # إطار مركزي للمعلومات
        center_frame = tk.Frame(self.info_frame, bg=self.colors["light"])
        center_frame.pack(expand=True)
        
        # إنشاء بطاقات معلومات جميلة
        cards_frame = tk.Frame(center_frame, bg=self.colors["light"])
        cards_frame.pack()
        
        # بطاقة مدة الدورة
        duration_card = self.create_info_card(
            cards_frame, 
            "مدة الدورة",
            "---",
            self.colors["primary"],
            0, 0
        )
        self.duration_var = duration_card
        
        # بطاقة فئة الدورة
        category_card = self.create_info_card(
            cards_frame,
            "فئة الدورة", 
            "---",
            self.colors["secondary"],
            0, 1
        )
        self.category_var = category_card
        
        # بطاقة الغياب بدون عذر
        without_excuse_card = self.create_info_card(
            cards_frame,
            "الحد الأقصى للغياب بدون عذر",
            "---",
            self.colors["danger"],
            1, 0
        )
        self.max_without_excuse_var = without_excuse_card
        
        # بطاقة الغياب بعذر
        with_excuse_card = self.create_info_card(
            cards_frame,
            "الحد الأقصى للغياب بعذر",
            "---",
            self.colors["warning"],
            1, 1
        )
        self.max_with_excuse_var = with_excuse_card
        
        # إطار قائمة المتدربين
        students_frame = tk.LabelFrame(
            self.window,
            text="قائمة المتدربين وحالة الغياب",
            font=self.fonts["subtitle"],
            bg=self.colors["light"],
            fg=self.colors["dark"],
            padx=10,
            pady=10
        )
        students_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # إنشاء Treeview
        columns = ("name", "rank", "without_excuse", "with_excuse", "total", "status", "percentage")
        
        tree_scroll = tk.Scrollbar(students_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.students_tree = ttk.Treeview(
            students_frame,
            columns=columns,
            show="tree headings",
            yscrollcommand=tree_scroll.set,
            style="Bold.Treeview"
        )
        self.students_tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.students_tree.yview)
        
        # تعريف الأعمدة
        self.students_tree.column("#0", width=150, anchor=tk.CENTER)
        self.students_tree.column("name", width=200, anchor=tk.CENTER)
        self.students_tree.column("rank", width=100, anchor=tk.CENTER)
        self.students_tree.column("without_excuse", width=120, anchor=tk.CENTER)
        self.students_tree.column("with_excuse", width=120, anchor=tk.CENTER)
        self.students_tree.column("total", width=100, anchor=tk.CENTER)
        self.students_tree.column("status", width=150, anchor=tk.CENTER)
        self.students_tree.column("percentage", width=100, anchor=tk.CENTER)
        
        # تعريف العناوين
        self.students_tree.heading("#0", text="رقم الهوية")
        self.students_tree.heading("name", text="اسم المتدرب")
        self.students_tree.heading("rank", text="الرتبة")
        self.students_tree.heading("without_excuse", text="غياب بدون عذر")
        self.students_tree.heading("with_excuse", text="غياب بعذر")
        self.students_tree.heading("total", text="المجموع")
        self.students_tree.heading("status", text="الحالة")
        self.students_tree.heading("percentage", text="النسبة %")
        
        # تعريف الألوان للحالات
        self.students_tree.tag_configure("safe", background="#e8f5e9")
        self.students_tree.tag_configure("warning", background="#fff8e1")
        self.students_tree.tag_configure("danger", background="#ffebee")
        self.students_tree.tag_configure("exceeded", background="#ff5252", foreground="white")
        
        # إطار الأزرار
        button_frame = tk.Frame(self.window, bg=self.colors["light"], pady=10)
        button_frame.pack(fill=tk.X, padx=10)
        
        tk.Button(
            button_frame,
            text="تحديث البيانات",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=15,
            pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.refresh_data
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="تصدير تقرير المتجاوزين",
            font=self.fonts["text_bold"],
            bg=self.colors["danger"],
            fg="white",
            padx=15,
            pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.export_exceeded_report
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="عرض التفاصيل",
            font=self.fonts["text_bold"],
            bg=self.colors["success"],
            fg="white",
            padx=15,
            pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.show_student_details
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15,
            pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=self.window.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
    def create_info_card(self, parent, title, value, color, row, col):
        """إنشاء بطاقة معلومات جميلة"""
        card_frame = tk.Frame(parent, bg="white", relief=tk.RAISED, bd=2)
        card_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # إطار العنوان
        title_frame = tk.Frame(card_frame, bg=color, height=40)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text=title,
            font=self.fonts["text_bold"],
            bg=color,
            fg="white"
        ).pack(expand=True)
        
        # إطار القيمة
        value_frame = tk.Frame(card_frame, bg="white", pady=20)
        value_frame.pack(fill=tk.BOTH, expand=True)
        
        value_var = tk.StringVar(value=value)
        value_label = tk.Label(
            value_frame,
            textvariable=value_var,
            font=("Tajawal", 18, "bold"),
            bg="white",
            fg=color
        )
        value_label.pack(expand=True)
        
        # تحديد الحد الأدنى للحجم
        card_frame.grid_configure(ipadx=40, ipady=20)
        
        return value_var
        
    def load_courses(self):
        """تحميل قائمة الدورات"""
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT DISTINCT ci.course_name 
            FROM course_info ci
            WHERE ci.end_date_system IS NOT NULL AND ci.end_date_system != ''
            ORDER BY ci.course_name
        """)
        courses = [row[0] for row in cursor.fetchall()]
        self.course_combo['values'] = courses
        
    def on_course_selected(self, event=None):
        """عند اختيار دورة"""
        course_name = self.course_var.get()
        if not course_name:
            return
            
        # الحصول على معلومات الدورة
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT start_day, start_month, start_year, 
                   end_day, end_month, end_year, end_date_system
            FROM course_info
            WHERE course_name = ?
        """, (course_name,))
        
        course_info = cursor.fetchone()
        if not course_info:
            messagebox.showwarning("تنبيه", "لم يتم العثور على معلومات الدورة")
            return
            
        # حساب مدة الدورة
        end_date_str = course_info[6]
        if end_date_str:
            end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
            # نفترض أن تاريخ البداية هو أول يوم تم تسجيل حضور فيه
            cursor.execute("""
                SELECT MIN(date) FROM attendance a
                JOIN trainees t ON a.national_id = t.national_id
                WHERE t.course = ?
            """, (course_name,))
            
            start_date_result = cursor.fetchone()
            if start_date_result and start_date_result[0]:
                start_date = datetime.datetime.strptime(start_date_result[0], "%Y-%m-%d").date()
                duration_days = (end_date - start_date).days + 1
                
                # تحديد فئة الدورة وحدود الغياب
                category, max_without, max_with = self.determine_course_category(duration_days)
                
                # تحديث عرض المعلومات في البطاقات
                self.duration_var.set(f"{duration_days} يوم")
                self.category_var.set(category)
                self.max_without_excuse_var.set(f"{max_without} أيام")
                self.max_with_excuse_var.set(f"{max_with} أيام")
                
                # تحميل بيانات المتدربين
                self.load_students_data(course_name, max_without, max_with)
                
    def determine_course_category(self, duration_days):
        """تحديد فئة الدورة وحدود الغياب المسموح - محدث حسب المعايير الجديدة"""
        if duration_days >= 85:  # 3 أشهر أو أكثر
            return "3 أشهر أو أكثر", 4, 8
        elif duration_days >= 55:  # شهرين
            return "شهرين", 2, 4
        elif duration_days >= 25:  # شهر
            return "شهر واحد", 1, 2
        elif duration_days >= 15:  # 3 أسابيع
            return "3 أسابيع", 1, 2
        else:  # أقل من 3 أسابيع
            return "أقل من 3 أسابيع", 1, 2
            
    def calculate_actual_absence_days(self, student_id, course_name):
        """حساب أيام الغياب الفعلية مع قاعدة الخميس-الأحد"""
        cursor = self.conn.cursor()
        
        # الحصول على جميع سجلات الغياب للمتدرب
        cursor.execute("""
            SELECT date, status, excuse_reason
            FROM attendance
            WHERE national_id = ? AND status IN ('غائب', 'غائب بعذر')
            ORDER BY date
        """, (student_id,))
        
        absence_records = cursor.fetchall()
        
        without_excuse_days = 0
        with_excuse_days = 0
        processed_dates = set()
        
        for record in absence_records:
            date_str, status, excuse_reason = record
            
            if date_str in processed_dates:
                continue
                
            processed_dates.add(date_str)
            
            # تحويل التاريخ
            absence_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            
            # حساب الأيام الأساسية
            if status == "غائب":
                days_to_add = 1
                
                # التحقق من قاعدة الخميس-الأحد
                weekday = absence_date.weekday()
                
                # إذا كان الغياب يوم الخميس (weekday = 3)
                if weekday == 3:
                    # التحقق من غياب يوم الأحد التالي
                    sunday_date = absence_date + datetime.timedelta(days=3)
                    sunday_str = sunday_date.strftime("%Y-%m-%d")
                    
                    cursor.execute("""
                        SELECT status FROM attendance
                        WHERE national_id = ? AND date = ? AND status = 'غائب'
                    """, (student_id, sunday_str))
                    
                    if cursor.fetchone():
                        # إذا غاب الخميس والأحد، نحسب 4 أيام
                        days_to_add = 4
                        processed_dates.add(sunday_str)
                        
                # إذا كان الغياب يوم الأحد (weekday = 6)
                elif weekday == 6:
                    # التحقق من غياب يوم الخميس السابق
                    thursday_date = absence_date - datetime.timedelta(days=3)
                    thursday_str = thursday_date.strftime("%Y-%m-%d")
                    
                    cursor.execute("""
                        SELECT status FROM attendance
                        WHERE national_id = ? AND date = ? AND status = 'غائب'
                    """, (student_id, thursday_str))
                    
                    if cursor.fetchone() and thursday_str not in processed_dates:
                        # إذا غاب الخميس والأحد ولم نحسب الخميس بعد
                        days_to_add = 4
                        processed_dates.add(thursday_str)
                
                without_excuse_days += days_to_add
                
            elif status == "غائب بعذر":
                with_excuse_days += 1
                
        return without_excuse_days, with_excuse_days
    
    def load_students_data(self, course_name, max_without, max_with):
        """تحميل بيانات المتدربين وحساب حالة الغياب - مع ترتيب المتجاوزين أولاً"""
        # مسح البيانات الحالية
        for item in self.students_tree.get_children():
            self.students_tree.delete(item)
            
        cursor = self.conn.cursor()
        
        # الحصول على جميع المتدربين في الدورة
        cursor.execute("""
            SELECT national_id, name, rank
            FROM trainees
            WHERE course = ? AND is_excluded = 0
            ORDER BY name
        """, (course_name,))
        
        students = cursor.fetchall()
        
        # قوائم لتخزين المتدربين حسب الحالة
        exceeded_students = []
        warning_students = []
        danger_students = []
        safe_students = []
        
        for student in students:
            student_id, name, rank = student
            
            # حساب أيام الغياب الفعلية
            without_excuse, with_excuse = self.calculate_actual_absence_days(student_id, course_name)
            
            total_absence = without_excuse + with_excuse
            
            # تحديد الحالة
            status = "آمن"
            tag = "safe"
            percentage = 0
            
            # حساب النسبة من الحد المسموح
            without_percentage = (without_excuse / max_without * 100) if max_without > 0 else 0
            with_percentage = (with_excuse / max_with * 100) if max_with > 0 else 0
            max_percentage = max(without_percentage, with_percentage)
            
            student_data = {
                'id': student_id,
                'name': name,
                'rank': rank,
                'without_excuse': without_excuse,
                'with_excuse': with_excuse,
                'total': total_absence,
                'max_without': max_without,
                'max_with': max_with,
                'percentage': max_percentage
            }
            
            if without_excuse > max_without or with_excuse > max_with:
                student_data['status'] = "تجاوز الحد المسموح"
                student_data['tag'] = "exceeded"
                exceeded_students.append(student_data)
            elif without_percentage >= 80 or with_percentage >= 80:
                student_data['status'] = "تحذير - اقتراب من الحد"
                student_data['tag'] = "warning"
                warning_students.append(student_data)
            elif without_percentage >= 50 or with_percentage >= 50:
                student_data['status'] = "انتباه"
                student_data['tag'] = "danger"
                danger_students.append(student_data)
            else:
                student_data['status'] = "آمن"
                student_data['tag'] = "safe"
                safe_students.append(student_data)
        
        # إدخال المتدربين بالترتيب: المتجاوزين أولاً
        all_students = exceeded_students + warning_students + danger_students + safe_students
        
        for student_data in all_students:
            self.students_tree.insert(
                "",
                tk.END,
                text=student_data['id'],
                values=(
                    student_data['name'],
                    student_data['rank'],
                    f"{student_data['without_excuse']} / {student_data['max_without']}",
                    f"{student_data['with_excuse']} / {student_data['max_with']}",
                    student_data['total'],
                    student_data['status'],
                    f"{student_data['percentage']:.1f}%"
                ),
                tags=(student_data['tag'],)
            )
            
        # عرض ملخص
        if len(exceeded_students) > 0:
            messagebox.showwarning(
                "تنبيه هام",
                f"يوجد {len(exceeded_students)} متدرب تجاوزوا الحد المسموح للغياب\n"
                f"و {len(warning_students)} متدرب في حالة تحذير"
            )
            
    def refresh_data(self):
        """تحديث البيانات"""
        if self.course_var.get():
            self.on_course_selected()
            messagebox.showinfo("تحديث", "تم تحديث البيانات بنجاح")
            
    def show_student_details(self):
        """عرض تفاصيل غياب المتدرب المحدد"""
        selection = self.students_tree.selection()
        if not selection:
            messagebox.showwarning("تنبيه", "الرجاء اختيار متدرب من القائمة")
            return
            
        student_id = self.students_tree.item(selection[0])['text']
        student_name = self.students_tree.item(selection[0])['values'][0]
        
        # إنشاء نافذة التفاصيل
        detail_window = tk.Toplevel(self.window)
        detail_window.title(f"تفاصيل غياب: {student_name}")
        detail_window.geometry("800x500")
        detail_window.configure(bg=self.colors["light"])
        
        # العنوان
        tk.Label(
            detail_window,
            text=f"سجل غياب المتدرب: {student_name}",
            font=self.fonts["title"],
            bg=self.colors["primary"],
            fg="white",
            padx=10,
            pady=10
        ).pack(fill=tk.X)
        
        # إطار التفاصيل
        detail_frame = tk.Frame(detail_window, bg=self.colors["light"], padx=10, pady=10)
        detail_frame.pack(fill=tk.BOTH, expand=True)
        
        # شجرة التفاصيل
        columns = ("date", "day", "status", "reason", "notes")
        detail_tree = ttk.Treeview(
            detail_frame,
            columns=columns,
            show="headings",
            height=15
        )
        detail_tree.pack(fill=tk.BOTH, expand=True)
        
        # تعريف الأعمدة
        detail_tree.column("date", width=100, anchor=tk.CENTER)
        detail_tree.column("day", width=100, anchor=tk.CENTER)
        detail_tree.column("status", width=100, anchor=tk.CENTER)
        detail_tree.column("reason", width=200, anchor=tk.CENTER)
        detail_tree.column("notes", width=200, anchor=tk.CENTER)
        
        # تعريف العناوين
        detail_tree.heading("date", text="التاريخ")
        detail_tree.heading("day", text="اليوم")
        detail_tree.heading("status", text="الحالة")
        detail_tree.heading("reason", text="السبب")
        detail_tree.heading("notes", text="ملاحظات")
        
        # تحميل سجلات الغياب
        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT date, status, excuse_reason
            FROM attendance
            WHERE national_id = ? AND status IN ('غائب', 'غائب بعذر')
            ORDER BY date DESC
        """, (student_id,))
        
        for record in cursor.fetchall():
            date_str, status, reason = record
            date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            day_name = self.get_arabic_day_name(date_obj)
            
            notes = ""
            # التحقق من قاعدة الخميس-الأحد
            if status == "غائب" and day_name in ["الخميس", "الأحد"]:
                if day_name == "الخميس":
                    sunday_date = date_obj + datetime.timedelta(days=3)
                    cursor.execute("""
                        SELECT status FROM attendance
                        WHERE national_id = ? AND date = ? AND status = 'غائب'
                    """, (student_id, sunday_date.strftime("%Y-%m-%d")))
                    if cursor.fetchone():
                        notes = "محسوب 4 أيام (خميس-أحد)"
                else:  # الأحد
                    thursday_date = date_obj - datetime.timedelta(days=3)
                    cursor.execute("""
                        SELECT status FROM attendance
                        WHERE national_id = ? AND date = ? AND status = 'غائب'
                    """, (student_id, thursday_date.strftime("%Y-%m-%d")))
                    if cursor.fetchone():
                        notes = "محسوب مع الخميس (4 أيام)"
                        
            detail_tree.insert(
                "",
                tk.END,
                values=(date_str, day_name, status, reason or "-", notes)
            )
            
        # زر الإغلاق
        tk.Button(
            detail_window,
            text="إغلاق",
            font=self.fonts["text_bold"],
            bg=self.colors["dark"],
            fg="white",
            padx=15,
            pady=5,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=detail_window.destroy
        ).pack(pady=10)
        
    def get_arabic_day_name(self, date_obj):
        """الحصول على اسم اليوم بالعربية"""
        days = {
            0: "الاثنين",
            1: "الثلاثاء", 
            2: "الأربعاء",
            3: "الخميس",
            4: "الجمعة",
            5: "السبت",
            6: "الأحد"
        }
        return days.get(date_obj.weekday(), "")
        
    def export_exceeded_report(self):
        """تصدير تقرير بالمتدربين المتجاوزين للحد المسموح"""
        import pandas as pd
        from tkinter import filedialog
        
        # جمع البيانات
        exceeded_students = []
        
        for child in self.students_tree.get_children():
            item = self.students_tree.item(child)
            if "exceeded" in item['tags']:
                values = item['values']
                exceeded_students.append({
                    'رقم الهوية': item['text'],
                    'الاسم': values[0],
                    'الرتبة': values[1],
                    'الغياب بدون عذر': values[2],
                    'الغياب بعذر': values[3],
                    'المجموع': values[4],
                    'الحالة': values[5],
                    'النسبة': values[6]
                })
                
        if not exceeded_students:
            messagebox.showinfo("معلومات", "لا يوجد متدربين متجاوزين للحد المسموح")
            return
            
        # تصدير إلى Excel
        export_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"المتجاوزين_للحد_المسموح_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if export_file:
            df = pd.DataFrame(exceeded_students)
            
            # إضافة معلومات الدورة
            course_name = self.course_var.get()
            duration = self.duration_var.get()
            category = self.category_var.get()
            
            with pd.ExcelWriter(export_file, engine='xlsxwriter') as writer:
                # معلومات الدورة
                info_df = pd.DataFrame({
                    'المعلومات': ['اسم الدورة', 'مدة الدورة', 'فئة الدورة', 'تاريخ التقرير'],
                    'القيمة': [course_name, duration, category, datetime.datetime.now().strftime('%Y-%m-%d')]
                })
                info_df.to_excel(writer, sheet_name='معلومات_الدورة', index=False)
                
                # قائمة المتجاوزين
                df.to_excel(writer, sheet_name='المتجاوزين', index=False)
                
                # تنسيق
                workbook = writer.book
                worksheet = writer.sheets['المتجاوزين']
                
                # تنسيق الرأس
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#FF5252',
                    'font_color': 'white',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                    
            messagebox.showinfo("نجاح", f"تم تصدير التقرير إلى:\n{export_file}")

# دالة لإضافة الأيقونة في التطبيق الرئيسي - محدثة بدون قواعد الاحتساب
def add_absence_monitoring_icon(self):
    """إضافة أيقونة نظام مراقبة الغياب في التطبيق الرئيسي"""
    # إضافة تبويب جديد في نافذة التطبيق الرئيسية
    if hasattr(self, 'tab_control'):
        self.absence_monitor_tab = tk.Frame(self.tab_control, bg=self.colors["light"])
        self.tab_control.add(self.absence_monitor_tab, text="مراقبة الغياب")
        
        # إطار للمحتوى
        content_frame = tk.Frame(self.absence_monitor_tab, bg=self.colors["light"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # عنوان
        tk.Label(
            content_frame,
            text="نظام مراقبة الغياب والإنذارات",
            font=self.fonts["title"],
            bg=self.colors["light"],
            fg=self.colors["primary"]
        ).pack(pady=20)
        
        # وصف
        tk.Label(
            content_frame,
            text="نظام متقدم لمراقبة غياب المتدربين وفقاً لتعليمات التدريب المستديمة\n"
                 "يحسب الحد المسموح للغياب حسب مدة الدورة ويُصدر إنذارات تلقائية",
            font=self.fonts["text"],
            bg=self.colors["light"],
            justify=tk.CENTER
        ).pack(pady=10)
        
        # زر فتح النظام
        tk.Button(
            content_frame,
            text="فتح نظام مراقبة الغياب",
            font=self.fonts["text_bold"],
            bg=self.colors["primary"],
            fg="white",
            padx=30,
            pady=15,
            bd=0,
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: AbsenceMonitoringSystem(self.root, self, self.colors, self.fonts)
        ).pack(pady=20)
