import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl import Workbook
import os
from datetime import datetime

class ColorVisionTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Тест на цветовосприятие")

        #Базы ответов для таблиц Рабкина

        # Нормальные ответы
        self.rabkin_normal_answers = {
            1: "96", 2: "О△", 3: "9", 4: "△", 5: "13", 6: "О△", 7: "96",
            8: "5", 9: "9", 10: "136", 11: "О△", 12: "12", 13: "О△", 14: "30",
            15: "О△", 16: "96", 17: "△О", 18: "=", 19: "95", 20: "О△", 21: "‖",
            22: "66", 23: "36", 24: "14", 25: "9", 26: "4", 27: "13",
            # Контрольная группа
            28: "14", 29: "2", 30: "О", 31: "2", 32: "9△", 33: "2", 34: "15",
            35: "5", 36: "О△", 37: "△", 38: "□", 39: "22", 40: "25", 41: "6",
            42: "Ничего", 43: "О△", 44: "О△", 45: "О△", 46: "60", 47: "=", 48: "Ничего"
        }

        # Ответы Протанопов (Пр)
        self.rabkin_protan_answers = {
            1: "96", 2: "О△", 3: "5", 4: "О", 5: "6", 6: "Ничего", 7: "96", 8: "5", 
            9: ["6", "8"], 10: ["68", "69"], 11: "△", 12: "Ничего", 13: "О", 
            14: "106", 15: ["△", "△□"], 16: "9", 17: "△", 18: "‖", 19: "5", 
            20: "Ничего", 21: "=", 22: "6", 23: "36", 24: "14", 25: "9", 26: "4", 27: "Ничего",
            # Контрольная группа        
            28: "14", 29: "2", 30: "Ничего", 31: "Ничего", 32: "9△", 33: "Ничего", 
            34: "Ничего", 35: "Ничего", 36: "Ничего", 37: "О", 38: "Ничего", 
            39: "Ничего", 40: "2", 41: "6", 42: "Ничего", 43: "О△",
            44: "Ничего", 45: "О△", 46: "60", 47: "‖", 48: "Ничего"
        }

        # Ответы Дейтеранопов (Де)
        self.rabkin_deutan_answers = {
            1: "96", 2: "О△", 3: "5", 4: "О", 5: "6", 6: "Ничего", 7: "6", 8: "5", 
            9: "9", 10: ["66", "69"], 11: ["△","О△"], 12: "12", 13: "△", 
            14: "16", 15: "△□", 16: "6", 17: "О", 18: "‖", 19: "5", 
            20: "Ничего", 21: "=", 22: "6", 23: "36", 24: "14", 25: "9", 26: "4", 27: "Ничего",
            # Контрольная группа
            28: "14", 29: "2", 30: "Ничего", 31: "Ничего", 32: "9△", 33: "Ничего", 
            34: "Ничего", 35: "Ничего", 36: "Ничего", 37: "О", 38: "Ничего", 
            39: "Ничего", 40: "2", 41: "6", 42: "Ничего", 43: "О△",
            44: "Ничего", 45: "О△", 46: "60", 47: "‖", 48: "Ничего"
        }
        
        self.protan_severity_map = {
            1: "+++", 2: "+++", 3: "---", 4: "---", 5: "---", 6: "+++", 7: "+++", 8: "+++",
            9: "---", 10: "---", 11: "---", 12: "---", 13: "---", 14: "---", 15: "---",
            16: "--+", 17: "--+", 18: "-++", 19: "+++", 20: "---", 21: "-++", 22: "---",
            23: "+++", 24: "+++", 25: "+++", 26: "+++", 27: "--+"
        }

        self.deutan_severity_map = {
            1: "+++", 2: "+++", 3: "---", 4: "---", 5: "---", 6: "+++", 7: "--+", 8: "+++",
            9: "+++", 10: "---", 11: "+++", 12: "+++", 13: "-++", 14: "---", 15: "---",
            16: "--+", 17: "--+", 18: "-++", 19: "--+", 20: "---", 21: "-++", 22: "---",
            23: "+++", 24: "+++", 25: "+++", 26: "+++", 27: "--+"
        }

        self.acquired_pathology_map = {
            1: "+", 2: "+", 3: "+", 4: "-", 5: "-", 6: "-", 7: "-",
            8: "-", 9: "+", 10: "-", 11: "+", 12: "-", 13: "+", 14: "-",
            15: "-", 16: "-", 17: "-", 18: "+", 19: "+", 20: "-", 21: "-",
            22: "-", 23: "-", 24: "-", 25: "-", 26: "-", 27: "-",
            # Контрольная группа
            28: "+", 29: "-", 30: "-", 31: "-", 32: "+", 33: "-", 34: "-",
            35: "-", 36: "-", 37: "-", 38: "-", 39: "-", 40: "-", 41: "-",
            42: "+", 43: "+", 44: "-", 45: "+", 46: "-", 47: "-", 48: "+"
        }
        self.total_acquired = sum(1 for v in self.acquired_pathology_map.values() if v == '-')
        
        #Базы ответов для таблиц Ишихара

        #Нормальные ответы
        self.ishihara_normal_answers = {
            1: "12", 2: "8", 3: "29", 4: "57", 5: "5", 6: "3", 7: "15",
            8: "74", 9: "97", 10: "45", 11: "5", 12: "73", 13: "Ничего", 14: "Ничего",
            15: "Ничего", 16: "26", 17: "42", 18: "35", 19: "96", 
            20: "Обе кривые", 21: "Обе кривые",
            22: "Ничего", 23: "О", 24: "О", 25: "О", 26: "О", 27: "О",
            28: "О", 29: "О", 30: "О", 31: "О", 32: "О", 33: "О", 34: "О",
            35: "О", 36: "О", 37: "О", 38: "О"
        }

        self.ishihara_type1_answers = {
            1: "12", 2: "3", 3: "70", 4: "35", 5: "2", 6: "5", 7: "17",
            8: "21", 9: "Ничего", 10: "Ничего", 11: "Ничего", 12: "Ничего", 
            13: "5", 14: "2", 15: "45", 16: "6", 17: "2", 18: "5", 19: "6", 
            20: "Нижняя кривая", 21: "Верхняя кривая", 22: "-", 23: "-", 
            24: "-", 25: "-", 26: "-", 27: "-", 28: "-", 29: "-", 30: "-", 
            31: "О", 32: "-", 33: "-", 34: "-", 35: "-", 36: "-", 37: "-", 38: "О"            
        }

        self.ishihara_type2_answers = {
            1: "12", 2: "3", 3: "70", 4: "35", 5: "2", 6: "5", 7: "17",
            8: "21", 9: "Ничего", 10: "Ничего", 11: "Ничего", 12: "Ничего", 
            13: "5", 14: "2", 15: "45", 16: "2", 17: "4", 18: "3", 19: "9", 
            20: "Верхняя кривая", 21: "Нижняя кривая",22: "-", 
            23: "-", 24: "-", 25: "-", 26: "-", 27: "-", 28: "-", 29: "-", 
            30: "-", 31: "О", 32: "-", 33: "-", 34: "-", 35: "-", 36: "-", 37: "-", 38: "О"            
        }

        self.user_answers = {}
        self.current_table_index = 1
        self.total_tables = 0  
        self.current_test = ""

        self.setup_ui()

    def save_results(self, data):
        filename = "color_vision_test_results.xlsx"
        headers = [
            "Timestamp", "Test Type", "Eye", "Diagnosis", "Severity",
            "Confidence", "Acquired Pathology (%)", "Normal Count",
            "Protan/Type1 Count", "Deutan/Type2 Count", "User Answers"
        ]

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            data.get("Test Type", ""),
            data.get("Eye", ""),
            data.get("Diagnosis", ""),
            data.get("Severity", "N/A"),
            data.get("Confidence", ""),
            data.get("Acquired Probability", "N/A"),
            data.get("Normal Count", ""),
            data.get("Type1 Count", ""),
            data.get("Type2 Count", ""),
            str(data.get("User Answers", {}))
        ]

        try:
            if not os.path.exists(filename):
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Results"
                sheet.append(headers)
            else:
                workbook = openpyxl.load_workbook(filename)
                sheet = workbook.active
            
            sheet.append(row)
            workbook.save(filename)
            return True
        except PermissionError:
            messagebox.showerror("Ошибка сохранения", f"Не удалось сохранить файл '{filename}'.\nПожалуйста, закройте файл, если он открыт, и попробуйте снова.")
            return False
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла непредвиденная ошибка при сохранении файла:\n{e}")
            return False

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(main_frame, text="Выберите тест:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.test_type = ttk.Combobox(main_frame, values=["Таблицы Рабкина", "Таблицы Ишихара"], state="readonly")
        self.test_type.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(main_frame, text="Выберите глаз:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.eye_selection = ttk.Combobox(main_frame, values=["OD (правый)", "OS (левый)"], state="readonly")
        self.eye_selection.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        self.start_button = ttk.Button(main_frame, text="Начать тест", command=self.start_test)
        self.start_button.grid(row=2, column=0, columnspan=2, pady=10)

        self.test_frame = ttk.LabelFrame(main_frame, text="Прохождение теста", padding="10")
        self.test_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        self.test_frame.grid_remove() 

    def start_test(self):
        self.current_test = self.test_type.get()
        if not self.current_test or not self.eye_selection.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите тест и глаз.")
            return
        
        if self.current_test == "Таблицы Рабкина":
            self.total_tables = len(self.rabkin_normal_answers)
        elif self.current_test == "Таблицы Ишихара":
            self.total_tables = len(self.ishihara_normal_answers)
        
        self.user_answers = {}
        self.current_table_index = 1
        self.show_test_widgets()

    def show_test_widgets(self):
        for widget in self.test_frame.winfo_children():
            widget.destroy()

        if self.current_table_index > self.total_tables:
            self.reset_test_ui()
            if self.current_test == "Таблицы Рабкина":
                self.analyze_rabkin_results()
            elif self.current_test == "Таблицы Ишихара":
                self.analyze_ishihara_results()
            return

        self.test_frame.grid()
        self.start_button.config(state=tk.DISABLED)
        self.test_frame.config(text=f"{self.current_test} | Глаз: {self.eye_selection.get()}")

        ttk.Label(
            self.test_frame,
            text=f"Таблица №{self.current_table_index} из {self.total_tables}:"
        ).grid(row=0, column=0, sticky=tk.W, pady=5)

        if self.current_test == "Таблицы Рабкина":
            options = ["", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "О", "△", "□", "=", "‖", "Ничего"]
        else:
            options = ["", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "-", "О", "Обе кривые", "Нижняя кривая", "Верхняя кривая"]

        self.answer_var1 = tk.StringVar()
        self.answer_entry1 = ttk.Combobox(self.test_frame, textvariable=self.answer_var1, values=options, width=14, state="readonly")
        self.answer_entry1.grid(row=0, column=1, padx=4, pady=5)

        self.answer_var2 = tk.StringVar()
        self.answer_entry2 = ttk.Combobox(self.test_frame, textvariable=self.answer_var2, values=options, width=14, state="readonly")
        self.answer_entry2.grid(row=0, column=2, padx=4, pady=5)

        self.answer_var3 = tk.StringVar()
        self.answer_entry3 = ttk.Combobox(self.test_frame, textvariable=self.answer_var3, values=options, width=14, state="readonly")
        self.answer_entry3.grid(row=0, column=3, padx=4, pady=5)

        next_button = ttk.Button(self.test_frame, text="Далее", command=self.record_answer)
        next_button.grid(row=1, column=0, columnspan=4, pady=10)
        self.root.bind('<Return>', lambda event: next_button.invoke())

    def record_answer(self):
        raw_values = [self.answer_var1.get(), self.answer_var2.get(), self.answer_var3.get()]
        parts = [v for v in raw_values if v != ""]

        if not parts:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите хотя бы один символ или вариант ответа.")
            return

        single_only = {"Ничего", "Обе кривые", "Нижняя кривая", "Верхняя кривая"}
        if any(p in single_only for p in parts) and len(parts) > 1:
            messagebox.showerror(
                "Ошибка",
                "Варианты «Ничего» и ответы с кривыми нельзя комбинировать с другими символами."
            )
            return

        self.user_answers[self.current_table_index] = "".join(parts)
        self.current_table_index += 1
        self.show_test_widgets()

    def reset_test_ui(self):
        self.root.unbind('<Return>')
        for widget in self.test_frame.winfo_children():
            widget.destroy()
        self.test_frame.grid_remove()
        self.start_button.config(state=tk.NORMAL)

    def analyze_ishihara_results(self):
        normal_score, type1_score, type2_score = 0, 0, 0

        for table_num, user_answer in self.user_answers.items():
            if user_answer == self.ishihara_normal_answers.get(table_num):
                normal_score += 1
            if user_answer == self.ishihara_type1_answers.get(table_num):
                type1_score += 1
            if user_answer == self.ishihara_type2_answers.get(table_num):
                type2_score += 1

        total_mistakes = self.total_tables - normal_score

        if normal_score >= self.total_tables * 0.8:
            diagnosis = "Нормальное цветовосприятие."
            confidence = (normal_score / self.total_tables) * 100
            confidence_str = f"Доля совпадений с нормой: {confidence:.2f}%"
        elif type1_score > type2_score + 1 and type1_score > 2:
            diagnosis = "Подозрение на нарушение цветовосприятия 1-го типа."
            consistency = (type1_score / total_mistakes) * 100 if total_mistakes > 0 else 0
            confidence_str = f"Согласованность с 1-м типом: {consistency:.2f}% (среди всех ошибок)"
        elif type2_score > type1_score + 1 and type2_score > 2:
            diagnosis = "Подозрение на нарушение цветовосприятия 2-го типа."
            consistency = (type2_score / total_mistakes) * 100 if total_mistakes > 0 else 0
            confidence_str = f"Согласованность с 2-м типом: {consistency:.2f}% (среди всех ошибок)"
        else:
            diagnosis = "Неопределенный результат. Обнаружены ошибки, характерные для разных типов нарушений."
            error_rate = (total_mistakes / self.total_tables) * 100
            confidence_str = f"Доля отклонений от нормы: {error_rate:.2f}%"

        results_to_save = {
            "Test Type": self.current_test,
            "Eye": self.eye_selection.get(),
            "Diagnosis": diagnosis,
            "Confidence": confidence_str,
            "Normal Count": normal_score,
            "Type1 Count": type1_score,
            "Type2 Count": type2_score,
            "User Answers": self.user_answers
        }
        saved = self.save_results(results_to_save)
        save_message = "\nРезультаты сохранены в файл." if saved else ""

        result_message = (
            f"Тест завершен!\n\n"
            f"Совпадений с нормой: {normal_score} из {self.total_tables}\n"
            f"Совпадений с типом 1: {type1_score}\n"
            f"Совпадений с типом 2: {type2_score}\n\n"
            f"ПРЕДВАРИТЕЛЬНОЕ ЗАКЛЮЧЕНИЕ:\n{diagnosis}\n{confidence_str}\n\n"
            "ВНИМАНИЕ: Данный результат не является медицинским диагнозом."
            f"{save_message}"
        )
        messagebox.showinfo("Результаты теста", result_message)
        self.eye_selection.set('')

    def estimate_rabkin_severity(self, severity_map):
        levels = [("Сильная (A)", 0), ("Средняя (B)", 1), ("Слабая (C)", 2)]
        scores = {}

        for label, idx in levels:
            matches = 0
            for table_num in range(3, 28):  # Таблицы 1-2 - демонстрационные
                code = severity_map.get(table_num)
                if not code or len(code) != 3:
                    continue
                is_normal = (self.user_answers.get(table_num) == self.rabkin_normal_answers.get(table_num))
                expected_normal = (code[idx] == "+")
                if is_normal == expected_normal:
                    matches += 1
            scores[label] = matches

        max_score = max(scores.values())
        best_levels = [label for label, score in scores.items() if score == max_score]
        if len(best_levels) == 1:
            return f"Предполагаемая степень: {best_levels[0]}"
        return f"Предполагаемая степень: Не определена однозначно ({' / '.join(best_levels)})"
    
    def analyze_rabkin_results(self):
        normal_matches, protan_error_count, deutan_error_count, acquired_error_count = 0, 0, 0, 0

        for table_num, user_answer in self.user_answers.items():
            if user_answer == self.rabkin_normal_answers.get(table_num):
                normal_matches += 1
                continue

            if self.acquired_pathology_map.get(table_num) == '-':
                acquired_error_count += 1

            protan_expected = self.rabkin_protan_answers.get(table_num)
            is_protan_error = (user_answer in protan_expected) if isinstance(protan_expected, list) else (user_answer == protan_expected)
            if is_protan_error:
                protan_error_count += 1

            deutan_expected = self.rabkin_deutan_answers.get(table_num)
            is_deutan_error = (user_answer in deutan_expected) if isinstance(deutan_expected, list) else (user_answer == deutan_expected)
            if is_deutan_error:
                deutan_error_count += 1

        total_tables = len(self.rabkin_normal_answers)
        total_mistakes = total_tables - normal_matches
        diagnosis, severity_str, confidence_str = "", "", ""

        if normal_matches >= (total_tables - 2):
            diagnosis = "Нормальное цветовосприятие."
            confidence = (normal_matches / total_tables) * 100
            confidence_str = f"Доля совпадений с нормой: {confidence:.2f}%"

        elif protan_error_count > deutan_error_count + 1 and protan_error_count > 2:
            diagnosis = "Подозрение на нарушение в красно-зеленом спектре (протанопия / протаномалия)."
            severity_str = self.estimate_rabkin_severity(self.protan_severity_map)
            consistency = (protan_error_count / total_mistakes) * 100 if total_mistakes > 0 else 0
            confidence_str = f"Согласованность с паттерном протанопии: {consistency:.2f}% (среди всех ошибок)"

        elif deutan_error_count > protan_error_count + 1 and deutan_error_count > 2:
            diagnosis = "Подозрение на нарушение в красно-зеленом спектре (дейтеранопия / дейтераномалия)."
            severity_str = self.estimate_rabkin_severity(self.deutan_severity_map)
            consistency = (deutan_error_count / total_mistakes) * 100 if total_mistakes > 0 else 0
            confidence_str = f"Согласованность с паттерном дейтеранопии: {consistency:.2f}% (среди всех ошибок)"

        else:
            diagnosis = "Неопределенный результат. Рекомендуется консультация специалиста."
            error_rate = (total_mistakes / total_tables) * 100
            confidence_str = f"Доля отклонений от нормы: {error_rate:.2f}%"

        acquired_probability = (acquired_error_count / self.total_acquired) * 100
        acquired_str = f"Ошибок на таблицах, чувствительных к приобретённым патологиям: {acquired_probability:.2f}%"

        results_to_save = {
            "Test Type": self.current_test,
            "Eye": self.eye_selection.get(),
            "Diagnosis": diagnosis,
            "Severity": severity_str.replace("Предполагаемая степень: ", "") if severity_str else "N/A",
            "Confidence": confidence_str,
            "Acquired Probability": f"{acquired_probability:.2f}%",
            "Normal Count": normal_matches,
            "Type1 Count": protan_error_count,
            "Type2 Count": deutan_error_count,
            "User Answers": self.user_answers
        }
        saved = self.save_results(results_to_save)
        save_message = "\nРезультаты сохранены в файл." if saved else ""

        result_message = (
            f"Тест завершен!\n\n"
            f"Правильных ответов: {normal_matches} из {total_tables}\n"
            f"Ошибок, характерных для протанопии: {protan_error_count}\n"
            f"Ошибок, характерных для дейтеранопии: {deutan_error_count}\n\n"
            f"ПРЕДВАРИТЕЛЬНОЕ ЗАКЛЮЧЕНИЕ:\n"
            f"{diagnosis}\n"
            f"{severity_str}\n\n"
            f"{confidence_str}\n\n"
            f"{acquired_str}\n\n"
            "ВНИМАНИЕ: Данный результат не является медицинским диагнозом."
            f"{save_message}"
        )

        messagebox.showinfo("Результаты теста", result_message)
        self.eye_selection.set('')

if __name__ == "__main__":
    root = tk.Tk()
    app = ColorVisionTestApp(root)
    root.mainloop()