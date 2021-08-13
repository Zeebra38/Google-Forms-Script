# Google-Forms-Script
Автоматические создание и заполнение .docx и .pdf по шаблону, с использованием данных Гугл Таблицы, подключенной к Форме

Триггер необходимо повесить на `Гугл Таблицу`, то есть на **Таблицу с ответами формы**, а не на саму форму.

## Структура документов:
- Отдельная папка
  - Форма `%Название формы%`
  - Гугл Таблица (создается автоматически после нажатия в форме `Создать таблицу`) `%Название формы% (Ответы)`
  - Шаблон документа Гугл Документы `%Название формы% (Шаблон)`
  - Папка с документами, прикрепляемыми к форме (генерируется автоматически после создания формы) `%Название формы% (File responses)`. К этой папке необходимо дать доступ по ссылке!
  - Папка с готовыми документами `%Название формы% Документы`

![image of documents structure](https://user-images.githubusercontent.com/68852325/128478721-ab8416f3-0035-45d8-811a-cd2edd3a85b5.png)

Перед началом работы **очень важно** создать именно такую структуру, во избежание ошибок.

## Структура таблицы с ответами:
* Первый столбец - `Отметка об исполнении`
* Потом идут автоматически сгенерированные столбцы с ответами на форму.
  * Поле для загрузки файлов должно содержать в себе `Приложение`
* Следующий столбец (без пропусков) надо назвать `Edit URL`
* После пропуска одно столбца необходимо именовать столбцы с людьми, которым на почту будет приходить подтверждение формы. Формат названия столбца `email/Комментарий`. Комментарий может содержать пробелы, но пробелов между `/` ни с одной стороны быть **не должно**. То есть например `mail@mail.ru/Очень длинный комментрий с пробелами`. Комментарий не должен содержать в себе `/`. Последний "одобритель" является адресатом. То есть ему приходит письмо с уже проверенной и одобренной служебной запиской.
  * Между столбцами "одобрителей" могут быть пробелы. После столбцов с "одобрителями" не должно быть других данных

## Как создать шаблон для документов
* Все переменные обрамляются кавычками `<<%Название переменной%>>`
  * Название должно быть точно такое же, как и в Форме, иначе не будут подставляться значения. **Регистр важен!**
  * Можно использовать форматирования текста, как в Гугл Документы. Подчеркивание, курсив, размер шрифта и т.д.
  * Не должно быть названий, которых нет в форме, они будут оставлены без изменений.
    * Есть несколько исключений. `День`, `Месяц`, `Год`, `Номер`, `Название приложения`  - служебно зарезервированные слова. 
      * `День` - число от 1 до 31, представляющее собой день месяца, в который будет создан документ по образцу
      * `Месяц` - представляет собой буквенное представление месяца
      * `Год` - представляет собой численное представление года
      * `Номер` - номер служебной записки. Определяется строкой в таблице, то есть нет документа с номером 1
      * `Название приложения` - названия всех файлов приложений. **Название пункта с загружаемыми файлами должно содержать в себе слово `Приложение`!**

## Настройки формы
Необходимо сделать следующие шаги:
- [x] Поставить галочку `Изменять ответы после отправки формы`
- [x] Выбрать максимальный размер всех загружаемых файлов
- [x] В настройках больше ничего по функционалу не трогать

Поля необходимо называть используя только буквы, цифры и пробелы. Специальные символы использовать нельзя, кроме `№` и `,`, '-', `.`.
Поле для загрузки файлов должно содержать в себе `Приложение`
Можно сделать "дублирующее" поле. Вместо текста будет использоваться ссылка. Вот, например
![image](https://user-images.githubusercontent.com/68852325/129035950-12016797-245c-4ebd-94e7-5746ab8a2562.png)

В таком случае необходимо скопировать первое название, а в конце добавить `вложение при необходимости`. В таком случае, если в ответе будет **ПУСТОЙ** первый вариант, но будет приложен документ, то в итоговый документ будет добавлена ссылка на Google Disk.

## Как добавить скрипт на новые таблицы.
1. Зайти в изначальный редактор скриптов. Например, он был прикреплен к таблице "Счет выставить (Ответы)". Необходимо ее открыть, выбрать `Инструменты`, выбрать `Редактор скриптов`. Откроется окно редактора на файле `Изменять здесь.gs`.
2. Найти переменную `spreadsheetsIDs`. Найти в ней закрывающую  **`**. Перед ней ввести Id таблицы и нажать Enter. Смотри example для примера.
  * Как узнать Id? ![image](https://user-images.githubusercontent.com/68852325/129038281-e14dc3a1-de32-44cf-a313-f46d5adfeb98.png) Id выделено. Все, что после d/ и перед /edit
3. Нажать сверху кнопку `Выполнить` ![image](https://user-images.githubusercontent.com/68852325/129041030-5772a9c0-ec56-49e0-9739-4e31ff199140.png)
