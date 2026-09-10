import random
import pandas as pd

# ---------- Данные ----------
first_names_male = [
    "Александр", "Сергей", "Дмитрий", "Андрей", "Алексей",
    "Максим", "Евгений", "Иван", "Михаил", "Николай"
]
first_names_female = [
    "Анна", "Елена", "Ольга", "Мария", "Татьяна",
    "Наталья", "Ирина", "Екатерина", "Светлана", "Юлия"
]

invariant_last_names = [
    "Шевченко", "Коваленко", "Ткаченко", "Долгих", "Черных",
    "Ковальчук", "Бондарь", "Мельник", "Сковорода"
]

regular_last_names = [
    "Иванов", "Петров", "Смирнов", "Кузнецов", "Соколов",
    "Попов", "Лебедев", "Козлов", "Новиков", "Морозов"
]

male_patronymics = [
    "Иванович", "Петрович", "Сергеевич", "Андреевич", "Алексеевич",
    "Дмитриевич", "Михайлович", "Николаевич", "Владимирович", "Павлович"
]
female_patronymics = [
    "Ивановна", "Петровна", "Сергеевна", "Андреевна", "Алексеевна",
    "Дмитриевна", "Михайловна", "Николаевна", "Владимировна", "Павловна"
]

# ---------- Транслитерация ----------
TRANSLIT_MAP = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e',
    'ё': 'e', 'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k',
    'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r',
    'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'kh', 'ц': 'ts',
    'ч': 'ch', 'ш': 'sh', 'щ': 'shch', 'ъ': '', 'ы': 'y', 'ь': '',
    'э': 'e', 'ю': 'yu', 'я': 'ya'
}

def transliterate(text):
    """Переводит русский текст в латиницу (транслит)."""
    result = []
    for char in text:
        lower = char.lower()
        if lower in TRANSLIT_MAP:
            translit = TRANSLIT_MAP[lower]
            if char.isupper() and translit:
                translit = translit.capitalize()
            result.append(translit)
        else:
            result.append(char)
    return ''.join(result)

# ---------- Генерация одной персоны ----------
def generate_person(gender=None):
    if gender is None:
        gender = random.choice(['male', 'female'])

    # Имя
    if gender == 'male':
        first = random.choice(first_names_male)
        patronymic = random.choice(male_patronymics)
    else:
        first = random.choice(first_names_female)
        patronymic = random.choice(female_patronymics)

    # Фамилия
    if random.random() < 0.4:
        last = random.choice(invariant_last_names)
    else:
        last = random.choice(regular_last_names)
        if gender == 'female':
            last += 'а'

    # Полное оригинальное ФИО
    original = f"{last} {first} {patronymic}"

    # Транслит
    translit = f"{transliterate(last)} {transliterate(first)} {transliterate(patronymic)}"

    return {
        'Фамилия': last,
        'Имя': first,
        'Отчество': patronymic,
        'Пол': 'М' if gender == 'male' else 'Ж',
        'ФИО': original,
        'Транслит': translit
    }

# ---------- Генерация n персон ----------
def generate_people(n):
    return [generate_person() for _ in range(n)]

# ---------- Основная программа ----------
if __name__ == "__main__":
    try:
        n = int(input("Введите количество человек: "))
        if n <= 0:
            print("Число должно быть положительным.")
        else:
            people = generate_people(n)
            df = pd.DataFrame(people)

            # Вывод первых строк в консоль
            print("\nПервые 5 записей:")
            #print(df.head().to_string(index=False))
            print(df)
            # Сохранение в CSV
            csv_file = "people.csv"
            df.to_csv(csv_file, index=False, encoding='utf-8-sig')
            print(f"\nДанные сохранены в {csv_file}")

            # Сохранение в Excel (если установлен openpyxl)
            try:
                excel_file = "people.xlsx"
                df.to_excel(excel_file, index=False, engine='openpyxl')
                print(f"Данные сохранены в {excel_file}")
            except ImportError:
                print("Для сохранения в Excel установите openpyxl: pip install openpyxl")
            except Exception as e:
                print(f"Не удалось сохранить в Excel: {e}")

    except ValueError:
        print("Ошибка: введите целое число.")