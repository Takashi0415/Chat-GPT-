#英単語テスト(仮)
# ファイルパス
file_path_1 = r"C:\Users\takas\OneDrive\金のフレーズ 問題 (1).docx"
file_path_2 = r"C:\Users\takas\OneDrive\ドキュメント\金のフレーズ 回答 (1).docx"

# ファイル1を読み込む
with open(file_path_1, 'rb') as file1:
    content1 = file1.read()

# ファイル2を読み込む
with open(file_path_2, 'rb') as file2:
    content2 = file2.read()

# content1およびcontent2には、それぞれファイルの内容が格納されます


from docx import Document

def separate_questions(doc):
    questions = []
    current_question = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:  # 空でない場合
            current_question.append(text)
        elif current_question:  # 空で、かつ現在の問題がある場合
            questions.append(current_question)
            current_question = []

    # 最後の問題が空でない場合に追加
    if current_question:
        questions.append(current_question)

    return questions

file_path_1 = r"C:\Users\takas\OneDrive\金のフレーズ 問題 (1).docx"

# .docx ファイルを開く
doc = Document(file_path_1)

# 問題文ごとに区分けする
questions = separate_questions(doc)


from docx import Document

def separate_words_and_meanings(doc):
    words_and_meanings = []
    current_word_and_meaning = []

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:  # 空でない場合
            current_word_and_meaning.append(text)
        elif current_word_and_meaning:  # 空で、かつ現在の英単語とその単語の意味がある場合
            words_and_meanings.append(current_word_and_meaning)
            current_word_and_meaning = []

    # 最後の英単語とその単語の意味が空でない場合に追加
    if current_word_and_meaning:
        words_and_meanings.append(current_word_and_meaning)

    return words_and_meanings

file_path_2 = r"C:\Users\takas\OneDrive\ドキュメント\金のフレーズ 回答 (1).docx"

# .docx ファイルを開く
doc_2 = Document(file_path_2)

# 英単語とその単語の意味ごとに区分けする
words_and_meanings = separate_words_and_meanings(doc_2)


import tkinter as tk
from tkinter import messagebox
import random
from docx import Document

# テストデータ
file_path_1 = r"C:\Users\takas\OneDrive\金のフレーズ 問題 (1).docx"
file_path_2 = r"C:\Users\takas\OneDrive\ドキュメント\金のフレーズ 回答 (1).docx"

doc_1 = Document(file_path_1)
questions_1 = separate_questions(doc_1)

doc_2 = Document(file_path_2)
answers_1 = separate_questions(doc_2)

# 手動で問題と答えを対応づける
question_answer_dict = {tuple(question): tuple(answer) for question, answer in zip(questions_1, answers_1)}

class QuizPage(tk.Frame):
    def __init__(self, master, questions, answers, load_start_page_callback, load_result_page_callback):
        super().__init__(master)
        self.master = master
        self.master.title("Quiz App")
        self.master.geometry("400x300")

        self.label_question = tk.Label(self, text="", font=("Helvetica", 14))
        self.label_question.pack(pady=20)

        self.entry_answer = tk.Entry(self, font=("Helvetica", 12))
        self.entry_answer.pack(pady=10)
        self.entry_answer.bind('<Return>', self.check_answer)

        self.button_submit = tk.Button(self, text="Submit", command=self.check_answer)
        self.button_submit.pack(pady=10)

        # クイズの追加部分
        self.total_questions = 10  # 修正: クイズの総数を1に設定
        self.current_question_number = 0
        self.correct_answers = 0
        self.chances = 1
        self.question_answer_dict = {tuple(question): tuple(answer) for question, answer in zip(questions, answers)}  # 修正
        self.questions = questions
        self.answers = answers

        self.load_start_page_callback = load_start_page_callback
        self.load_result_page_callback = load_result_page_callback

        self.load_question()

    def load_question(self):
        global current_question, current_answer

        if self.current_question_number < self.total_questions:
            self.current_question_number += 1
            random.shuffle(self.questions)  # 修正: 問題をシャッフル
            current_question = tuple(self.questions[0])  # 修正: 質問をタプルに変換
            current_answer = self.question_answer_dict[current_question]  # 修正
            
            self.label_question.config(text=f"Question {self.current_question_number}/{self.total_questions}:\n"
                                         f"Japanese: {current_question[0]}\nEnglish: {current_question[1]}")
            self.entry_answer.delete(0, tk.END)
            self.chances = 1
        else:
            self.show_results()

    def check_answer(self, event=None):
        global current_question, current_answer

        user_answer = self.entry_answer.get().strip().lower()
        correct_word = current_answer[0].strip().lower()  # 修正: 答えの取得方法を変更

        if user_answer == correct_word:
            messagebox.showinfo("Correct", f"Correct answer: \n{current_answer[0]}\n{current_answer[1]}")
            self.correct_answers += 1
            self.show_answer_and_result()

        else:
            if self.chances < 3:
                messagebox.showerror("Incorrect", "Incorrect answer. Try again.")
                self.entry_answer.delete(0, tk.END)
                self.chances += 1
            else:
                messagebox.showinfo("Incorrect", f"Incorrect answer. Correct answer: \n{current_answer[0]}\n{current_answer[1]}")
                self.show_answer_and_result()

    def show_answer_and_result(self):
        if self.current_question_number == self.total_questions:
            self.show_results()
        else:
            self.load_question()

    def show_results(self):
        self.load_result_page_callback(self.correct_answers, self.total_questions)  # 修正: コールバック関数を呼ぶ
        self.current_question_number = 0
        self.correct_answers = 0
        self.load_question()


class StartPage(tk.Frame):
    def __init__(self, master, load_quiz_page_callback):
        tk.Frame.__init__(self, master)
        self.master = master
        self.master.title("English Word Quiz")
        self.master.geometry("400x300")

        self.label_title = tk.Label(self, text="English Word Quiz", font=("Helvetica", 18))
        self.label_title.pack(pady=20)

        self.start_button = tk.Button(self, text="Start Quiz", command=load_quiz_page_callback)
        self.start_button.pack()


class ResultPage(tk.Frame):
    def __init__(self, master, correct_answers, total_questions, restart_quiz_callback):
        super().__init__(master)
        self.master = master
        self.master.title("Result Page")
        self.master.geometry("400x300")

        self.label_result = tk.Label(self, text=f"You answered {correct_answers} out of {total_questions} questions correctly.")
        self.label_result.pack(pady=20)

        self.restart_button = tk.Button(self, text="Restart Quiz", command=restart_quiz_callback)
        self.restart_button.pack(pady=10)


class QuizApp:
    def __init__(self, master):
        self.master = master
        self.start_page = StartPage(self.master, self.load_quiz_page)
        self.quiz_page = QuizPage(self.master, questions_1, answers_1, self.load_start_page, self.load_result_page)
        self.result_page = None

    def load_start_page(self):
        self.quiz_page.pack_forget()
        self.start_page.pack(expand=True, fill="both")

    def load_quiz_page(self):
        self.start_page.pack_forget()
        self.quiz_page.pack(expand=True, fill="both")

    def load_result_page(self, correct_answers, total_questions):
        if self.quiz_page:
            self.quiz_page.pack_forget()
        self.result_page = ResultPage(self.master, correct_answers, total_questions, self.restart_quiz)
        self.result_page.pack(expand=True, fill="both")

    def restart_quiz(self):
        if self.result_page:
            self.result_page.pack_forget()
        self.start_page.pack(expand=True, fill="both")


if __name__ == "__main__":
    root = tk.Tk()
    app = QuizApp(root)
    app.start_page.pack(expand=True, fill="both")
    root.mainloop()
