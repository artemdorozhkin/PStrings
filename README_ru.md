# PStrings

[eng](README.md) | rus

PStrings - это легкая библиотека VBA, которая предоставляет различные функции манипулирования строками для упрощения общих задач в программировании на Visual Basic for Applications (VBA).

## Оглавление

- [Установка](#installation)
- [Использование](#usage)
- [Функции](#functions)
  - [FromatBytes](#formatbytes)
  - [Wrap](#wrap)
  - [IndexOf](#indexof)
  - [LastIndexOf](#lastindexof)
  - [CharCodeAt](#charcodeat)
  - [CharAt](#charat)
  - [Slice](#slice)
  - [StartsWith](#startswith)
  - [EndsWith](#endswith)
  - [Join](#join)
  - [Trim](#trim)
  - [JoinNonEmpty](#joinnonempty)
  - [FString](#fstring)
  - [IsNullString](#isnullstring)
  - [FormatString](#formatstring)
  - [InString](#instring)
  - [IsEqual](#isequals)

## Установка

Чтобы использовать PStrings в своем проекте VBA, выполните следующие действия:

1. Загрузите файл модуля PStrings из раздела [releases](https://github.com/example/PStrings/releases) этого репозитория.
2. Импортируйте модуль PStrings в свой проект VBA:
   - Откройте рабочую книгу Excel или проект VBA.
   - Нажмите `Alt + F11`, чтобы открыть редактор VBA.
   - Перейдите в меню `Файл > Импорт файла` и выберите загруженный файл модуля PStrings.
3. После импорта вы можете начать использовать функции PStrings в своем коде VBA.

## Использование

Вот как можно использовать PStrings в коде VBA:

````vb
' Пример использования функций PStrings
Sub ExampleUsage()
Dim text As String
text = "Hello, {0}!"
' Пример использования функции FString
Debug.Print FString(text, "world") ' Вывод: Hello, world!
End Sub

## Функции

### FormatBytes

```vb
Public Function FormatBytes(ByVal Amount As Long) As String

````

#### Описание:

Форматирует строку, представляющую сумму, указывая выход в килобайтах, мегабайтах или гигабайтах.

#### Параметры:

- `Amount`: Сумма в байтах.

#### Пример:

```vb
Debug.Print FormatBytes(1023) ' 1023 байта
Debug.Print FormatBytes(1024) ' 1,00 КБ
Debug.Print FormatBytes(1048576) ' 1,00 МБ
Debug.Print FormatBytes(1073741824) ' 1.00 GB
```

#### Возвращает:

Отформатированная строка, представляющая объем с соответствующей единицей измерения (байт, КБ, МБ, ГБ).

### Wrap

```vb
Public Function Wrap(ByVal Text As String, ByVal Wrapper As String) As String
```

#### Описание:

Функция `Wrap()` заворачивает вводимый текст в указанную обертку и возвращает новую строку.

#### Параметры:

- `Text`: Вводимый текст.
- `Wrapper`: Строка-обертка.

#### Пример:

```vb
Debug.Print Wrap("ABC", Chr(34)) ' "ABC"
Debug.Print Wrap("ABC", "ABC") ' ABCABCABC
```

#### Возвращает:

Строка, представляющая входной текст, обернутый в указанную обертку.

### IndexOf

```vb
Public Function IndexOf(ByVal Text As String, ByVal Char As String) As Integer
```

#### Описание:

Возвращает индекс первого появления указанного символа в заданной текстовой строке.

#### Параметры:

- `Text`: Текстовая строка для поиска.
- `Char`: Символ, который необходимо найти в текстовой строке.

#### Возвращает:

Индекс первого появления указанного символа в текстовой строке. Если символ не найден, возвращается -1.

#### Замечания:

Эта функция выполняет итерацию по каждому символу в текстовой строке и возвращает индекс первого вхождения указанного символа.

#### Пример:

```vb
Debug.Print IndexOf("hello", "e") ' Вывод: 2
Debug.Print IndexOf("hello", "z") ' Вывод: -1
```

### LastIndexOf

```vb
Public Function LastIndexOf(ByVal Text As String, ByVal Char As String) As Integer
```

#### Описание:

Возвращает индекс последнего появления указанного символа в заданной текстовой строке.

#### Параметры:

- `Text`: Текстовая строка для поиска.
- `Char`: Символ, который нужно найти в текстовой строке.

#### Возвращает:

Индекс последнего появления указанного символа в текстовой строке. Если символ не найден, возвращается -1.

#### Замечания:

Эта функция выполняет итерацию по каждому символу в текстовой строке в обратном порядке и возвращает индекс последнего вхождения указанного символа.

#### Пример:

```vb
Debug.Print LastIndexOf("hello", "l") ' Вывод: 4
Debug.Print LastIndexOf("hello", "z") ' Вывод: -1
```

### CharCodeAt

```vb
Public Function CharCodeAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As Integer
```

#### Описание:

Функция `CharCodeAt()` возвращает значение Unicode от 0 до 65535, представляющее код символа ASCII на основе указанного индекса.

#### Параметры:

- `Text`: Входная строка.
- `Index`: Индекс от 0 до Len(Str) - 1. Если индекс не указан, по умолчанию используется первый символ.

#### Возвращает:

Значение Unicode, представляющее код символа ASCII в указанном индексе. Если указана недопустимая строка, функция возвращает -1.

#### Пример:

```vb
Debug.Print CharCodeAt("ABC", 0) ' Вывод: 65
```

### CharAt

```vb
Public Function CharAt(ByVal Text As String, Optional ByVal Index As Integer = 0) As String
```

#### Описание:

Функция `CharAt()` возвращает новую строку, состоящую из одного символа, извлеченного из указанной индексной позиции во входной строке.

#### Параметры:

- `Text`: Входная строка.
- `Index`: Индекс от 0 до Len(Str) - 1. Если индекс не указан, по умолчанию используется первый символ.

#### Пример:

В следующем примере доступ к символам осуществляется в различных позициях строки "Brave new world":

```vb
Dim AnyString As String: AnyString = "Brave new world"
Debug.Print "Character at index 0 is " & CharAt(AnyString)
' Без указания индекса по умолчанию принимает значение 0.
Debug.Print "Character at index 0 is " & CharAt(AnyString, 0)
Debug.Print "Символ с индексом 1 - это " & CharAt(AnyString, 1)
Debug.Print "Символ с индексом 2 - это " & CharAt(AnyString, 2)
Debug.Print "Символ с индексом 3 - это " & CharAt(AnyString, 3)
Debug.Print "Символ с индексом 4 - это " & CharAt(AnyString, 4)
Debug.Print "Символ с индексом 999 - это " & CharAt(AnyString, 999)
```

Эти строки отображают следующее:

```vb
Символ с индексом 0 - это 'B'
Символ с индексом 0 - это 'B'
Символ под индексом 1 - 'r'
Символ под индексом 2 - 'a'
Символ под индексом 3 - 'v'
Символ с индексом 4 - 'e'
Символ с индексом 999 - ''
```

### Slice

```vb
Public Function Slice(ByVal Text As String, ByVal StartIndex As Integer, Optional ByVal EndIndex As Integer = -1) As String
```

#### Описание:

Возвращает подстроку заданного текста и конкатенирует ее в новую строку, исключая указанный конечный индекс.

#### Параметры:

- `Text`: Входная строка.
- `StartIndex`: Индекс первого символа, который нужно включить в результирующую подстроку.
- `EndIndex`: Индекс символа, следующего сразу за концом нужной подстроки. По умолчанию равен -1, что означает конец строки.

#### Пример:

Следующий пример демонстрирует функцию `Slice()` для создания новой подстроки:

```vb
Dim Text1 As String: Text1 = "Наступило утро". ' Длина Text1 равна 23.
Dim Text2 As String: Text2 = Slice(Text1, 1, 8)
Dim Text3 As String: Text3 = Slice(Text1, 4, -2)
Dim Text4 As String: Text4 = Slice(Text1, 12)
Dim Text5 As String: Text5 = Slice(Text1, 30)
Debug.Print Text2 ' "он умер"
Debug.Print Text3 ' "утро наступило"
Debug.Print Text4 ' "наступило".
Debug.Print Text5 ' ""
```

Следующий пример демонстрирует работу функции `Slice()` с начальным индексом по умолчанию:

```vb
Dim Text1 As String: Text1 = "Наступило утро".
Slice(Text1, -3) ' "нас".
Slice(Text1, -3, -1) ' "нас"
Slice(Text1, 0, -1) ' "Наступило утро"
Slice(Text1, 4, -1) ' "утро наступило"
```

В этом примере нарезка начинается с 11-го символа и заканчивается 16-м символом:

```vb
Slice(Text1, -11, 16) ' "is u"
```

В этих примерах нарезка начинается с 5-го символа и заканчивается 1-м символом:

```vb
Slice(Text1, -5, -1) ' "n us"
```

### StartsWith

```vb
Public Function StartsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
```

#### Описание:

Проверяет, совпадает ли начало текста с указанным `выражением`.

#### Параметры:

- `Текст`: Проверяемый текст.
- `Выражение`: Значение, которое необходимо найти.
- `Compare`: Метод сравнения. Перечисление `VbCompareMethod`. По умолчанию используется `vbBinaryCompare`.

#### Возвращает:

- `True`, если начало текста совпадает с указанным выражением; иначе `False`.

#### Пример:

Следующий пример возвращает `True`, потому что слово "Check" начинается с "che" и выбран метод сравнения `vbTextCompare`:

```vb
Debug.Print StartsWith("Check", "che", vbTextCompare)
```

Следующий пример возвращает `False`, потому что выбран метод сравнения по умолчанию `vbBinaryCompare`:

```vb
Debug.Print StartsWith("Check", "che")
```

### EndsWith

```vb
Public Function EndsWith(ByVal Text As String, ByVal Expression As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
```

#### Описание:

Проверяет, совпадает ли конец текста с указанным `выражением`.

#### Параметры:

- `Текст`: Проверяемый текст.
- `Выражение`: Значение, которое необходимо найти.
- `Compare`: Метод сравнения. Перечисление `VbCompareMethod`. По умолчанию используется `vbBinaryCompare`.

#### Возвращает:

- `True`, если конец текста совпадает с указанным выражением; иначе `False`.

#### Пример:

Следующий пример возвращает `True`, потому что слово "Check" заканчивается на "ECK" и выбран метод сравнения `vbTextCompare`:

```vb
Debug.Print EndsWith("Check", "ECK", vbTextCompare)
```

Следующий пример возвращает `False`, потому что выбран метод сравнения по умолчанию `vbBinaryCompare`:

```vb
Debug.Print EndsWith("Check", "ECK")
```

### Присоединяйтесь

```vb
Public Function Join(ByVal Delimiter As String, ByVal ParamArray Values() As String) As String
```

#### Описание:

Конкатенирует несколько строк с указанным разделителем.

#### Параметры:

- `Delimiter`: Разделитель.
- `Значения`: Строки для конкатенации.

#### Возвращает:

Конкатенированная строка с указанным разделителем.

#### Пример:

```vb
Dim Value1 As String: Value1 = "Объединенная строка"
Dim Value2 As String: Value2 = "из двух строк"
Debug.Print Join(" ", Value1, Value2) ' Объединенная строка из двух строк
```

### Trim

```vb
Public Function Trim(ByVal Text As String) As String
```

#### Описание:

Удаляет из строки ведущие и завершающие пробелы.

#### Параметры:

- `Text`: Строка с лидирующими и отстающими пробелами.

#### Пример:

```vb
Dim Text As String: Text = " Строка с пробелами "
Debug.Print ">" & Trim(Text) & "<" ' >Строка с пробелами<
```

### JoinNonEmpty

```vb
Public Function JoinNonEmpty(ByVal Data As Variant, Optional ByVal Delimiter As String = ", ") As String
```

#### Описание:

Конкатенирует элементы заданного массива, исключая пустые значения (`vbNullString` или `Empty`), с указанным разделителем.

#### Параметры:

- `Data`: Массив данных.
- `Delimiter`: Разделитель. По умолчанию используется запятая и пробел.

#### Возвращает:

Конкатенированная строка с непустыми значениями, разделенными указанным разделителем.

#### Пример:

```vb
Dim DataWithEmpty As Variant
DataWithEmpty = Array("Value1", Empty, "Value2", "")
Debug.Print JoinNonEmpty(DataWithEmpty, ", ") ' Value1, Value2
```

### FString

```vb
Public Function FString(ByVal Text As String, ParamArray Values() As Variant) As String
```

#### Описание:

Заменяет местоположения в тексте на соответствующие значения из массива `Values`.

#### Замечания:

Для интерполяции следует использовать местодержатели в виде {0}, {1} и т. д., соответствующие индексу значения в массиве.
Функция не обрабатывает экранирующие последовательности, такие как: `\n`, `\t`, `\r`.

#### Пример:

```vb
Dim Text as String
Text = "Пример использования функции {0}!"
Dim FuncName as String
FuncName = "FString"
Debug.Print FString(Text, FuncName) ' Пример использования функции FString!
```

### IsEmptyString

```vb
Public Function IsEmptyString(ByVal Expression As String) As Boolean
```

#### Описание:

Проверяет, является ли заданное значение `String` пустым.

#### Параметры:

- `Expression`: Значение для проверки.

#### Возвращает:

`True`, если значение пустое.

#### Пример:

```vb
Dim str As String
str = ""
Debug.Print IsEmptyString(str) ' True
str = "Hello"
Debug.Print IsEmptyString(str) ' False
```

### FormatString

```vb
Public Function FormatString(ByVal Text As String, ByVal ParamArray Values() As Variant) As String
```

#### Описание:

Форматирует местозаполнители в тексте соответствующими значениями из массива `Values`.

#### Замечания:

Для интерполяции используйте местодержатели в следующем формате:

- `%s` для значений String
- `%t` для значений даты
- `%d` для числовых значений

#### Пример:

```vb
Dim Text as String
Text = "Пример использования функции %s!"
Dim FuncName as String
FuncName = "FormatString"
Debug.Print FormatString(Text, FuncName) ' Результат: "Пример использования функции FormatString!"
```

#### Параметры:

- `Text`: Текст с заполнителями.

#### Возвращает:

Отформатированный текст.

### InString

```vb
Public Function InString(ByVal Text As String, ByVal ParamArray Values() As Variant) As Boolean
```

#### Описание:

Проверяет, присутствует ли какое-либо из `Значений` в строке `Текст`.

#### Параметры:

- `Текст`: Текст для поиска.
- `Значения`: Значения для поиска.

#### Возвращает:

Возвращает `True`, если в тексте найдено любое из указанных значений.

### IsEqual

```vb
Public Function IsEqual(ByVal Text1 As String, ByVal Text2 As String, Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
```

#### Описание:

Сравнивает равенство `Text1` и `Text2`.

#### Параметры:

- `Text1`: Первая строка.
- `Text2`: Вторая строка.
- `Compare`: Метод сравнения. По умолчанию используется `vbTextCompare`.

#### Возвращает:

Возвращает `True`, если строки равны.
