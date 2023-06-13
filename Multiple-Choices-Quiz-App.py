
import pandas as pd
import tkinter as tk
from tkinter import messagebox


def load_questions(filename):
    df = pd.read_excel(filename)
    questions = []
    for _, row in df.iterrows():
        question = {
            'question': row['Question'],
            'choices': [f"A. {row['A']}", f"B. {row['B']}", f"C. {row['C']}", f"D. {row['D']}", f"E. {row['E']}"],
            'answer': row['Correct']
        }
        questions.append(question)
    return questions


def display_question():
    question = questions[question_index]
    question_text = question['question']
    choices = question['choices']

    # Update GUI elements with new question data
    question_label.config(text=f"Question {question_index+1}: {question_text}")
    for i, choice in enumerate(choices):
        choice_buttons[i].config(text=choice)
    marks_label.config(text=f"Marks: {marks[question_index]}")
    correct_label.config(text="")


def check_answer():
    selected_answer = selected.get()
    if selected_answer == '1':
        c = 'A'
    elif selected_answer == '2':
        c = 'B'
    elif selected_answer == '3':
        c = 'C'
    elif selected_answer == '4':
        c = 'D'
    else:
        c = 'E'

    correct_answer = questions[question_index]['answer']
    print('Selected:', selected_answer)
    print('Correct:', correct_answer)
    if c == correct_answer:
        marks[question_index] = 1
        result_label.config(text="Correct!", fg="green")
       # messagebox.showinfo("Result", "Correct!")
    else:
        marks[question_index] = 0
        result_label.config(text="Incorrect!", fg="red")
       # messagebox.showinfo("Result", "Incorrect!")

    correct_label.config(text=f"Correct answer: {correct_answer}")


def next_question():
    global question_index
    question_index += 1
    if question_index < len(questions):
        display_question()
        selected.set(None)
        result_label.config(text="")
        correct_label.config(text="")
        back_button.config(state=tk.NORMAL)
        first_button.config(state=tk.NORMAL)
    else:
        # End of questions
        question_label.config(text="End of questions.")
        for choice_button in choice_buttons:
            choice_button.config(state=tk.DISABLED)
        next_button.config(state=tk.DISABLED)
        last_button.config(state=tk.NORMAL)


def previous_question():
    global question_index
    question_index -= 1
    if question_index >= 0:
        display_question()
        selected.set(None)
        result_label.config(text="")
        correct_label.config(text="")
        next_button.config(state=tk.NORMAL)
        last_button.config(state=tk.NORMAL)
    else:
        # First question
        back_button.config(state=tk.DISABLED)
        first_button.config(state=tk.DISABLED)


def first_question():
    global question_index
    question_index = 0
    display_question()
    selected.set(None)
    result_label.config(text="")
    correct_label.config(text="")
    back_button.config(state=tk.DISABLED)
    first_button.config(state=tk.DISABLED)
    next_button.config(state=tk.NORMAL)
    last_button.config(state=tk.NORMAL)


def last_question():
    global question_index
    question_index = len(questions) - 1
    display_question()
    selected.set(None)
    result_label.config(text="")
    correct_label.config(text="")
    next_button.config(state=tk.DISABLED)
    last_button.config(state=tk.DISABLED)
    back_button.config(state=tk.NORMAL)
    first_button.config(state=tk.NORMAL)


def exit_exam():
    if messagebox.askokcancel("Exit", "Are you sure you want to exit the exam?"):
        window.destroy()


# Load questions from Excel file write the correct path to the excel file
questions = load_questions(r"***:\****\********.xlsx")

# GUI setup
window = tk.Tk()
window.title("Multiple Choice Quiz")
window.geometry(f"{1680}x1050")

# Font settings
font_size = 12
font_family = "Arial"

question_label = tk.Label(window, text="", wraplength=800, font=(font_family, font_size, "bold"))
question_label.pack()

selected = tk.StringVar()
choice_buttons = []
for i in range(5):
    choice_button = tk.Radiobutton(window, text="", variable=selected, value=str(i + 1), font=(font_family, font_size))
    choice_button.pack()
    choice_buttons.append(choice_button)

submit_button = tk.Button(window, text="Submit", command=check_answer,width='10',height='2')
submit_button.pack()

result_label = tk.Label(window, text="")
result_label.pack()

correct_label = tk.Label(window, text="")
correct_label.pack()

marks_label = tk.Label(window, text="",font=(font_family, font_size))
marks_label.pack()

# Next and Back buttons
button_frame = tk.Frame(window)
button_frame.pack()

first_button = tk.Button(button_frame, text="First", command=first_question,width='10',height='2')
first_button.pack(side=tk.LEFT)

back_button = tk.Button(button_frame, text="Back", state=tk.DISABLED, command=previous_question,width='10',height='2')
back_button.pack(side=tk.LEFT)

next_button = tk.Button(button_frame, text="Next", command=next_question,width='10',height='2')
next_button.pack(side=tk.LEFT)

last_button = tk.Button(button_frame, text="Last", command=last_question,width='10',height='2')
last_button.pack(side=tk.LEFT)

# Exit button
exit_button = tk.Button(window, text="Exit", command=exit_exam,width='10',height='2')
exit_button.pack()

# Start with the first question
question_index = 0
marks = [0] * len(questions)

display_question()
window.mainloop()
