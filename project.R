# Установка библиотек
library(dplyr)
library(magrittr)
library(readr)
library(tidyr)
library(stringr)
library(lubridate)
library(openxlsx)
library(ggplot2)

# На основе это функции, выполняеться чтение табличных данных из папки
read_data <- function(file) {
  if (file.exists(file)) {
    read_csv(file)
  } else {
    message(paste("Данные отсутствуют:", file))
    NULL
  }
}

# Применяю функцию к данным
financial_data <- read_data("financial_data.csv")
prolongations <- read_data("prolongations.csv")

# Просмотр первых 10 строк
head(financial_data, 10)
head(prolongations, 10)


# Чтобы было удобне работать с табличными данными я преобразовал табличные данные из широкого в узкий формат
financial_long <- function(df){
  df %>% 
    pivot_longer(
    cols = -c("id", "Причина дубля", "Account"),
    names_to = "month",
    values_to = "shipment"
  )
}

# Применяю функцию к датафрейму
financial_transform <- financial_long(financial_data)


# Трансформирую колонки "shipment, month"
# По условию заменяю "в ноль" на последнюю ненулевую отгрузку
# Также заменяю все "стоп", "end" на NA
financial_replacement <- function(df) { 
  df %>%
    arrange(id, month) %>%  
    group_by(id) %>%
    mutate(
      month_raw = str_to_title(str_trim(month)),
      year = str_extract(month_raw, "\\d{4}"),
      month = str_extract(month_raw, "^[А-Яа-я]+"),  
      shipment = str_replace_all(shipment, ",", "."),
      shipment = str_replace_all(shipment, "\\s+", ""),
      shipment = ifelse(shipment == "в ноль", "0", shipment),
      shipment = ifelse(shipment %in% c("стоп", "end", ""), NA, shipment),
      shipment = str_extract(shipment, "[0-9 .,]+"),
      shipment = as.numeric(shipment),
      shipment = ifelse(shipment == 0 & all(shipment == 0), lag(shipment, default = 0), shipment),
      shipment = ifelse(is.na(shipment), 0, shipment)
    ) %>%
    select(-month_raw) %>% 
    ungroup()
}

# Применяю функцию к датафрейму
financial_finaly <- financial_replacement(financial_transform)

# Выполняеться левое обьединение с таблицей пролонгаций по полю id
# Если поле AM (тоисть ФИО менеджера) из таблицы "prolongations" равно NA, заменяю на поле из таблицы "financial_finaly", в противном случае оставляем "AM - ФИО"
# Группирую по "ID, AM, month"
# Произвожу вычисление по сумме отгрузки
financiall_full <- financial_finaly %>%
  left_join(prolongations %>% select(id, AM), by = "id") %>%
  mutate(AM = ifelse(is.na(AM), Account, AM)) %>% 
  group_by(id, AM, month, year) %>%
  summarise(shipment = sum(shipment, na.rm = TRUE)) %>%
  ungroup()

# Данная функция на вход берёт общий датафрейм, месяц, год
# Расчитывает 2-а коэффициента согласно условию 
calc_coef <- function(df, month_current, year_current) {
  months_order <- c("Январь","Февраль","Март","Апрель","Май",
                    "Июнь","Июль","Август","Сентябрь","Октябрь",
                    "Ноябрь","Декабрь")
  
  idx <- match(month_current, months_order)
  month_prev1 <- if(idx > 1) months_order[idx-1] else NA
  month_prev2 <- if(idx > 2) months_order[idx-2] else NA
  
  df %>%
    filter(year == as.character(year_current)) %>%
    group_by(AM) %>%
    summarise(
      coef1 = {
        prev1_vals <- shipment[month == month_prev1]
        current_vals <- shipment[month == month_current & id %in% id[month == month_prev1]]
        if(length(prev1_vals)==0 || sum(prev1_vals, na.rm = TRUE)==0) NA
        else sum(current_vals, na.rm = TRUE)/sum(prev1_vals, na.rm = TRUE)
      },
      coef2 = {
        prev2_vals <- shipment[month == month_prev2]
        ids_prev1 <- id[month == month_prev1 & shipment > 0]
        prev2_no_prev1_vals <- shipment[month == month_prev2 & !id %in% ids_prev1]
        current_vals2 <- shipment[month == month_current & id %in% id[month == month_prev2 & !id %in% ids_prev1]]
        if(length(prev2_no_prev1_vals)==0 || sum(prev2_no_prev1_vals, na.rm = TRUE)==0) NA
        else sum(current_vals2, na.rm = TRUE)/sum(prev2_no_prev1_vals, na.rm = TRUE)
      },
      .groups = "drop"
    )
}

# Данная функция на вход принимает датайфрейм, месяц, год и выполняет расчёт для каждой группы согласно условию использую функцию "calc_coef"
# А именно: Необходимо расчитать коэффициенты пролонгации для каждого менеджера и для всего отдела в целом:
# a.	за каждый месяц
# b.	за год
calc_all <- function(df, months_order, years) {
  all_months <- lapply(years, function(y) {
    lapply(months_order, function(m) {
      tmp <- calc_coef(df, m, y)  
      tmp$month <- m
      tmp$year <- y
      tmp
    }) %>% bind_rows()
  }) %>% bind_rows()
  
  coef_per_manager_year <- all_months %>%
    group_by(AM, year) %>%
    summarise(
      coef1_year = mean(coef1, na.rm = TRUE),
      coef2_year = mean(coef2, na.rm = TRUE)
    )
  coef_department_month <- all_months %>%
    group_by(year, month) %>%
    summarise(
      coef1 = mean(coef1, na.rm = TRUE),
      coef2 = mean(coef2, na.rm = TRUE)
    )
  coef_department_year <- all_months %>%
    group_by(year) %>%
    summarise(
      coef1_year = mean(coef1, na.rm = TRUE),
      coef2_year = mean(coef2, na.rm = TRUE)
    )
  
  return(list(
    per_month = all_months,
    per_manager_year = coef_per_manager_year,
    department_month = coef_department_month,
    department_year = coef_department_year
  ))
}

# Вызов функции
result <- calc_all(financiall_full, months_order, 2022:2023)

# Просмотр результата с помощью цикла
for (name in names(result)) {
  cat("=== ", name, " ===\n")
  if(!is.null(result[[name]])) {
    print(head(result[[name]]))
  } else {
    message("Список пуст")
  }
  cat("\n")
}

# Данный график визуально показывает, в какие месяцы коэффициент был высоким, а в какие низким или отсутствовал
# Подготовка данных по отделу за каждый месяц и визуализация
df <- result$department_month %>%
  mutate(
    month_num = factor(match(month, c("Январь","Февраль","Март","Апрель","Май",
                                      "Июнь","Июль","Август","Сентябрь","Октябрь",
                                      "Ноябрь","Декабрь")),
                       levels = 1:12,
                       labels = c("Январь","Февраль","Март","Апрель","Май",
                                  "Июнь","Июль","Август","Сентябрь","Октябрь",
                                  "Ноябрь","Декабрь")),
    coef1 = ifelse(is.nan(coef1), 0, coef1),
    coef2 = ifelse(is.nan(coef2), 0, coef2)
  ) %>% 
  ggplot(aes(x = month_num, y = coef1, fill = coef1)) +
  geom_bar(stat = "identity", width = 1) +
  coord_polar(start = 0) +
  theme_minimal() +
  labs(title = "Радиальные столбцы коэффициента №1 для всего отдела по каждому месяцу", x = "", y = "") +
  theme(axis.text.y = element_blank(),
        axis.ticks = element_blank(),
        panel.grid = element_blank())

# Создание папки
image_dir <- "images"
if (!dir.exists(image_dir)) dir.create(image_dir)

# Сохраняю изображение в папку
ggsave(filename = file.path(image_dir, "plot.png"), plot = df, width = 8, height = 6)

# Выгрузка аналитического отчета в Excel
wb <- createWorkbook()
addWorksheet(wb, "Менеджеры по месяцам")
writeData(wb, "Менеджеры по месяцам", result$per_month)
addWorksheet(wb, "Менеджеры по годам")
writeData(wb, "Менеджеры по годам", result$per_manager_year)
addWorksheet(wb, "Отдел по месяцам")
writeData(wb, "Отдел по месяцам", result$department_month)
addWorksheet(wb, "Отдел по годам")
writeData(wb, "Отдел по годам", result$department_year)
saveWorkbook(wb, "prolongation_report.xlsx", overwrite = TRUE)