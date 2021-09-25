require 'pry'
require 'roo'

xlsx = Roo::Spreadsheet.open('./sheet.xlsx')
sheet = xlsx.sheet(0)


def delete_last_elements(values, k)
  (0..k).each do
    values.pop()
  end
end

# Данные про обучающий план
# Ну тут понятно, вытягиваем просто основные данные таблички
branch_of_knowledge    = sheet.cell('M', 12)
speciality             = sheet.cell('H', 13)
qualification          = sheet.cell('F', 14)
type_of_education      = sheet.cell('AQ', 12)
period                 = sheet.cell('AQ', 13)
educational_background = sheet.cell('AO', 14)

# ////////////////////////////
# Стягивание графика обучающего процесса
first_course  = sheet.row(21)
second_course = sheet.row(22)
third_course  = sheet.row(23)
fourth_course = sheet.row(24)

# Удаление лишних елеметов из недели. Обязательно оставить, так как идёт скрап целой строки из таблички
# В итоге мы получаем все елементы из таблицы на каждый семестр в массиве
delete_last_elements(first_course, 2)
delete_last_elements(second_course, 2)
delete_last_elements(third_course, 2)
delete_last_elements(fourth_course, 2)

# ///////////////////////////

# Данные про бюджет времени
#Верхние значения можно захардкодить все, потому их нет смысла переносить в код, но на всякий случай указал в description_budget
description_budget   = [sheet.cell('C', 31), sheet.cell('F', 31), sheet.cell('I', 31),
                        sheet.cell('L', 31), sheet.cell('O', 31), sheet.cell('R', 31)]
first_course_budget  = [sheet.cell('C', 33), sheet.cell('F', 33), sheet.cell('I', 33),
                        sheet.cell('L', 33), sheet.cell('O', 33), sheet.cell('R', 33)]
second_course_budget = [sheet.cell('C', 34), sheet.cell('F', 34), sheet.cell('I', 34),
                        sheet.cell('L', 34), sheet.cell('O', 34), sheet.cell('R', 34)]
third_course_budget  = [sheet.cell('C', 35), sheet.cell('F', 35), sheet.cell('I', 35),
                        sheet.cell('L', 35), sheet.cell('O', 35), sheet.cell('R', 35)]
fourth_course_budget = [sheet.cell('C', 36), sheet.cell('F', 36), sheet.cell('I', 36),
                        sheet.cell('L', 36), sheet.cell('O', 36), sheet.cell('R', 36)]

# ///////////////////////////
# Практика. Тут та же самая история, как и в прошлый раз

educational_practice    = [sheet.cell('AA', 32), sheet.cell('AJ', 32), sheet.cell('AL', 32)]
introductory_practice   = [sheet.cell('AA', 33), sheet.cell('AJ', 33), sheet.cell('AL', 33)]
bzhd_practice           = [sheet.cell('AA', 34), sheet.cell('AJ', 34), sheet.cell('AL', 34)]
safety_health_course    = [sheet.cell('AA', 35), sheet.cell('AJ', 35), sheet.cell('AL', 35)]
internship_practice     = [sheet.cell('AA', 36), sheet.cell('AJ', 36), sheet.cell('AL', 36)]
comprehensive_training  = [sheet.cell('AA', 37), sheet.cell('AJ', 37), sheet.cell('AL', 37)]
cientific_practice      = [sheet.cell('AA', 38), sheet.cell('AJ', 38), sheet.cell('AL', 38)]
prediploma_practice     = [sheet.cell('AA', 39), sheet.cell('AJ', 39), sheet.cell('AL', 39)]

# /////////////////////////
# Пидсумкова аттестация
english_exam   = [sheet.cell('AQ', 32), sheet.cell('AZ', 32)]
thesis_defence = [sheet.cell('AQ', 33), sheet.cell('AJ', 33)]
