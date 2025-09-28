import psycopg2
import pandas as pd

# Подключение к PostgreSQL
conn = psycopg2.connect(
    host="localhost",       
    port="5432",            
    database="dv_project",  
    user="postgres",        
    password="0708"   
)

# --- SQL запросы ---

queries = {
    # 1. Средний итоговый балл по курсам
    "avg_total_by_course": """
        SELECT g.course, AVG(g.total) AS avg_total
        FROM dv.grades_raw g
        GROUP BY g.course
        ORDER BY avg_total DESC;
    """,

    # 2. Количество студентов в каждой программе (ОП)
    "total_students_per_program": """
        SELECT s.op, COUNT(*) AS total_students
        FROM dv.students_raw s
        GROUP BY s.op
        ORDER BY total_students DESC;
    """,

    # 3. Посещаемость по группам
    "avg_attendance_by_group": """
        SELECT s.gruppa, AVG(a.attendance) AS avg_attendance
        FROM dv.students_raw s
        JOIN dv.attendance a ON s.email = a.email
        GROUP BY s.gruppa
        ORDER BY avg_attendance DESC;
    """,

    # 4. Минимальные и максимальные баллы по курсам
    "min_max_scores_by_course": """
        SELECT g.course, MIN(g.total) AS min_total, MAX(g.total) AS max_total
        FROM dv.grades_raw g
        GROUP BY g.course
        ORDER BY g.course;
    """,

    # 5. Количество студентов, выбравших каждый элективный курс
    "students_per_elective": """
        SELECT e.course_id, COUNT(e.email) AS num_students
        FROM dv.enrollment_raw e
        GROUP BY e.course_id
        ORDER BY num_students DESC;
    """,

    # 6. Распределение студентов по степеням
    "students_by_degree": """
        SELECT s.stepen, COUNT(*) AS total_students
        FROM dv.students_raw s
        GROUP BY s.stepen;
    """,

    # 7. Список студентов с посещаемостью меньше 70%
    "low_attendance_students": """
        SELECT s.fio, s.email, a.attendance
        FROM dv.students_raw s
        JOIN dv.attendance a ON s.email = a.email
        WHERE a.attendance < 70
        ORDER BY a.attendance ASC;
    """,

    # 8. (повтор) Средний итоговый балл по курсам
    "avg_total_duplicate": """
        SELECT g.course, AVG(g.total) AS avg_total
        FROM dv.grades_raw g
        GROUP BY g.course
        ORDER BY avg_total DESC;
    """,

    # 9. JOIN студентов и их оценок
    "students_with_grades": """
        SELECT s.fio, s.email, g.course, g.total
        FROM dv.students_raw s
        JOIN dv.grades_raw g
        ON s.email = g.email
        LIMIT 20;
    """,

    # 10. Студенты группы SE-2423
    "students_in_group_SE2423": """
        SELECT * 
        FROM dv.students_raw
        WHERE gruppa = 'SE-2423'
        ORDER BY fio;
    """
}

# --- Выполнение запросов ---
results = {}
for name, query in queries.items():
    df = pd.read_sql_query(query, conn)
    results[name] = df

    print(f"\n--- {name} ---")
    print(df.head(10))  # выводим первые 10 строк

    # Сохраняем в CSV
    df.to_csv(f"{name}.csv", index=False)

# Сохраняем всё в один Excel
with pd.ExcelWriter("results.xlsx") as writer:
    for name, df in results.items():
        df.to_excel(writer, sheet_name=name[:31], index=False)  # sheet_name <= 31 символ

# Закрыть соединение
conn.close()
print("\nDone! Все результаты сохранены в CSV и Excel.")
