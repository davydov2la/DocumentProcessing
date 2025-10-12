# DocumentProcessing

Библиотека для **обезличивания документов Word и SolidWorks**.  
Поддерживает как **OpenXML**, так и **COM Interop** обработку документов, а также пакетную обработку.

---

## 🚀 Возможности

- Обработка документов Microsoft Word (`.docx`, `.doc`)
- Обработка файлов SolidWorks (`.slddrw`, `.sldprt`, `.sldasm`)
- Пакетная (batch) обработка множества файлов
- Удаление идентификационных кодов и персональных данных
- Двухпроходная обработка с удалением кодов организаций
- Логирование действий через `ILogger`
- Настраиваемые стратегии поиска и замены

---

## ⚙️ Сборка

1. Установите [.NET SDK 9.0](https://dotnet.microsoft.com/en-us/download)
2. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/davydov2la/DocumentProcessing.git
3. Перейдите в каталог проекта и соберите решение
   ```bash
   cd DocumentProcessing
   dotnet build -c Release
   
---

## 💻 Использование

Пример простой анонимизации документа:
```csharp
using DocumentProcessingLibrary.Facade;

var anonymizer = new DocumentAnonymizer(visible: false, useOpenXml: true);
var result = anonymizer.AnonymizeDocument(
    inputFilePath: @"C:\docs\report.docx",
    outputDirectory: @"C:\docs\out");

Console.WriteLine($"Успех: {result.Success}, найдено: {result.MatchesFound}");
```

---

## 🧩 Структура проекта

DocumentProcessing/  
├──DocumentProcessing/ – основная библиотека (.NET)  
├── DocumentProcessingConsole/ – консольное приложение для тестов  
├── .gitignore  
├── DocumentProcessing.sln  
├── README.md  
└── LICENSE

---

## ⚠️ Ограничения

- COM-обработка (Word Interop, SolidWorks) поддерживается только на Windows.
- Для серверной обработки рекомендуется использовать OpenXML-режим.
- SolidWorks API требует установленный SolidWorks.

---

## 🧾 Лицензия
Проект распространяется под лицензией MIT.  
См. файл [LICENSE](LICENSE) для подробностей.



