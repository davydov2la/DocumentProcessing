using DocumentProcessing.Core.Interfaces;
using DocumentProcessing.Core.Strategies.Replacement;
using DocumentProcessing.Core.Strategies.Search;
using DocumentProcessing.Documents.Interfaces;
using DocumentProcessing.Facade;
using DocumentProcessing.Logging;
using DocumentProcessing.Processing.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace DocumentProcessingConsole
{
    public class ManualConfigurationExample
    {
        private class BatchFileResult
        {
            public string FilePath { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public bool Success { get; set; }
            public int MatchesFound { get; set; }
            public int MatchesProcessed { get; set; }
            public int ExtractedCodes { get; set; }
            public int CodesRemoved { get; set; }
            public List<string> Warnings { get; set; } = new();
            public List<string> Errors { get; set; } = new();
        }
        
        public static void ProcessBatch()
        {
            string inputDirectory = @"/Users/paveldavydov/RiderProjects/DocumentProcessing/ProcessingTest/Input";
            string outputDirectory = @"/Users/paveldavydov/RiderProjects/DocumentProcessing/ProcessingTest/Output";
    
            Console.WriteLine("=== ПАКЕТНАЯ ОБРАБОТКА С ПОЛНОЙ НАСТРОЙКОЙ ===\n");
    
            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Information);
                builder.AddFilter("DocumentProcessing", LogLevel.Debug);
                builder.AddFilter("Microsoft", LogLevel.Warning);
            });
    
            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<DocumentAnonymizer>>();
    
            Console.WriteLine("Шаг 1: Логгирование настроено\n");
    
            var designationSearchStrategy = new RegexSearchStrategy(
                "CustomDesignations",
                new RegexPattern(
                    "DesignationFullFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+\.(?:[0-9]{2,2}\.){2,}[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?"
                ),
                new RegexPattern(
                    "DesignationShortFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+-[А-Я0-9-]+\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                ),
                new RegexPattern(
                    "DesignationMinimalFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9]+\.[0-9]{2,2}\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                )
            );
    
            var nameSearchStrategy = new RegexSearchStrategy(
                "PersonNames",
                new RegexPattern("SurnameFirst", @"[А-Я][а-я]+\s[А-Я]\.\s?[А-Я]\."),
                new RegexPattern("InitialsFirst", @"[А-Я]\.\s?[А-Я]\.\s?[А-Я][а-я]+")
            );
    
            var codeExtractionStrategy = new OrganizationCodeRemovalStrategy();
            var nameRemovalStrategy = new RemoveReplacementStrategy();
    
            Console.WriteLine("Шаг 2: Стратегии поиска созданы\n");
    
            var firstPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>
                {
                    designationSearchStrategy,
                    nameSearchStrategy
                },
                ReplacementStrategy = new CompositeReplacementStrategy(
                    "FirstPassReplacement",
                    match => match.MatchType.Contains("Designation"),
                    codeExtractionStrategy,
                    nameRemovalStrategy
                ),
                Options = new ProcessingOptions
                {
                    ProcessProperties = true,
                    ProcessTextBoxes = true,
                    ProcessNotes = true,
                    ProcessHeaders = true,
                    ProcessFooters = true,
                    MinMatchLength = 5,
                    CaseSensitive = false
                },
                Logger = logger
            };
    
            var secondPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>(),
                ReplacementStrategy = new RemoveReplacementStrategy(),
                Options = new ProcessingOptions
                {
                    ProcessProperties = false,
                    ProcessTextBoxes = true,
                    ProcessNotes = true,
                    ProcessHeaders = true,
                    ProcessFooters = true,
                    MinMatchLength = 1,
                    CaseSensitive = true
                },
                Logger = logger
            };
    
            var twoPassConfig = new TwoPassProcessingConfiguration
            {
                FirstPassConfiguration = firstPassConfig,
                SecondPassConfiguration = secondPassConfig,
                CodeExtractionStrategy = codeExtractionStrategy
            };
    
            Console.WriteLine("Шаг 3: Конфигурации созданы\n");
    
            var files = Directory.GetFiles(inputDirectory, "*.docx", SearchOption.TopDirectoryOnly);
    
            if (files.Length == 0)
            {
                Console.WriteLine("Файлы не найдены в директории: " + inputDirectory);
                return;
            }
    
            Console.WriteLine($"Шаг 4: Найдено файлов для обработки: {files.Length}\n");
            Console.WriteLine("=".PadRight(60, '='));
    
            using (var anonymizer = new DocumentAnonymizer(
                visible: false,
                useOpenXml: true,
                logger: logger))
            {
                var batchResults = new List<BatchFileResult>();
                int currentFile = 0;
    
                foreach (var filePath in files)
                {
                    currentFile++;
                    var fileName = Path.GetFileName(filePath);
    
                    Console.WriteLine($"\n[{currentFile}/{files.Length}] Обработка: {fileName}");
                    Console.WriteLine("-".PadRight(60, '-'));
    
                    var fileResult = new BatchFileResult
                    {
                        FilePath = filePath,
                        FileName = fileName
                    };
    
                    try
                    {
                        var processingRequest = new DocumentProcessingRequest
                        {
                            InputFilePath = filePath,
                            OutputDirectory = outputDirectory,
                            Configuration = firstPassConfig,
                            ExportOptions = new ExportOptions
                            {
                                ExportToPdf = false,
                                SaveModified = true,
                                PdfFileName = null,
                                Quality = PdfQuality.HighQuality
                            },
                            PreserveOriginal = true
                        };
    
                        var result = new ProcessingResult();
                        var resultAwareLogger = new ResultAwareLogger(logger, result);
                        firstPassConfig.Logger = resultAwareLogger;
                        secondPassConfig.Logger = resultAwareLogger;
    
                        using (var factory = new DocumentProcessing.Documents.Factories.DocumentProcessorFactory(
                            visible: false,
                            useOpenXml: true,
                            logger: resultAwareLogger))
                        {
                            var processor = factory.CreateProcessor(filePath);
    
                            if (processor is ITwoPassDocumentProcessor twoPassProcessor)
                            {
                                result = twoPassProcessor.ProcessTwoPass(processingRequest, twoPassConfig);
                            }
                            else
                            {
                                result = processor.Process(processingRequest);
                            }
                        }
    
                        fileResult.Success = result.Success;
                        fileResult.MatchesFound = result.MatchesFound;
                        fileResult.MatchesProcessed = result.MatchesProcessed;
                        fileResult.Warnings = result.Warnings;
                        fileResult.Errors = result.Errors;
    
                        var extractedCodes = codeExtractionStrategy.GetExtractedCodes();
                        fileResult.ExtractedCodes = extractedCodes.Count;
    
                        if (result.Metadata.TryGetValue("CodesRemoved", out var codesRemoved))
                        {
                            fileResult.CodesRemoved = (int)codesRemoved;
                        }
    
                        codeExtractionStrategy.ClearExtractedCodes();
    
                        if (result.Success)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine($"✓ Успешно: найдено {result.MatchesFound}, обработано {result.MatchesProcessed}");
                            Console.ResetColor();
    
                            if (extractedCodes.Count > 0)
                            {
                                Console.WriteLine($"Извлечено кодов: {extractedCodes.Count}");
                            }
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine($"Ошибка: {string.Join(", ", result.Errors)}");
                            Console.ResetColor();
                        }
    
                        if (result.Warnings.Any())
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"Предупреждений: {result.Warnings.Count}");
                            Console.ResetColor();
                        }
                    }
                    catch (Exception ex)
                    {
                        fileResult.Success = false;
                        fileResult.Errors.Add(ex.Message);
    
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Исключение: {ex.Message}");
                        Console.ResetColor();
                    }
    
                    batchResults.Add(fileResult);
                }
    
                Console.WriteLine("\n" + "=".PadRight(60, '='));
                Console.WriteLine("\n=== ИТОГИ ПАКЕТНОЙ ОБРАБОТКИ ===\n");
    
                var successful = batchResults.Count(r => r.Success);
                var failed = batchResults.Count(r => !r.Success);
                var totalMatches = batchResults.Sum(r => r.MatchesFound);
                var totalProcessed = batchResults.Sum(r => r.MatchesProcessed);
                var totalCodes = batchResults.Sum(r => r.ExtractedCodes);
                var totalCodesRemoved = batchResults.Sum(r => r.CodesRemoved);
    
                Console.WriteLine($"Статистика:");
                Console.WriteLine($"Всего файлов: {files.Length}");
                Console.WriteLine($"Успешно: {successful}");
                Console.WriteLine($"Ошибок: {failed}");
                Console.WriteLine($"Всего найдено совпадений: {totalMatches}");
                Console.WriteLine($"Всего обработано: {totalProcessed}");
                Console.WriteLine($"Извлечено кодов: {totalCodes}");
                Console.WriteLine($"Удалено отдельных кодов: {totalCodesRemoved}\n");
    
                Console.WriteLine("Детали по файлам:");
                foreach (var fileResult in batchResults)
                {
                    if (fileResult.Success)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write("OK");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write("Not OK");
                        Console.ResetColor();
                    }
    
                    Console.WriteLine($"{fileResult.FileName}: {fileResult.MatchesProcessed} обработано");
    
                    if (fileResult.Warnings.Any())
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"{fileResult.Warnings.Count} предупреждений");
                        Console.ResetColor();
                    }
    
                    if (fileResult.Errors.Any())
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"Ошибки: {string.Join(", ", fileResult.Errors)}");
                        Console.ResetColor();
                    }
                }
    
                Console.WriteLine($"\nРезультаты сохранены в: {outputDirectory}");
            }
    
            Console.WriteLine("\n=== ПАКЕТНАЯ ОБРАБОТКА ЗАВЕРШЕНА ===");
        }
        
        /// <summary>
        /// ГЛАВНЫЙ ПРИМЕР: Полная ручная настройка с двухпроходной обработкой + логгирование
        /// </summary>
        public static void FullManualConfiguration()
        {
            string inputFile = @"/Users/paveldavydov/RiderProjects/DocumentProcessing/ProcessingTest/ЛР1 — копия.docx";
            string outputDir = @"/Users/paveldavydov/RiderProjects/DocumentProcessing/ProcessingTest/Output";

            Console.WriteLine("=== ПОЛНАЯ РУЧНАЯ НАСТРОЙКА БИБЛИОТЕКИ С ЛОГГИРОВАНИЕМ ===\n");

            // ========================================================================
            // ШАГ 0: Настройка логгирования
            // ========================================================================

            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Debug); // Для детального вывода
                builder.AddFilter("Microsoft", LogLevel.Warning); // Фильтруем Microsoft логи
            });

            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<DocumentAnonymizer>>();

            Console.WriteLine("Шаг 0: Логгирование настроено");
            Console.WriteLine("Console logger активирован");
            Console.WriteLine("Уровень: Debug\n");

            // ========================================================================
            // ШАГ 1: Создаем стратегии поиска
            // ========================================================================

            var designationSearchStrategy = new RegexSearchStrategy(
                "CustomDesignations",
                new RegexPattern(
                    "DesignationFullFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+\.(?:[0-9]{2,2}\.){2,}[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?"
                ),
                new RegexPattern(
                    "DesignationShortFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+-[А-Я0-9-]+\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                ),
                new RegexPattern(
                    "DesignationMinimalFormat",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9]+\.[0-9]{2,2}\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                )
            );

            var nameSearchStrategy = new RegexSearchStrategy(
                "PersonNames",
                new RegexPattern("SurnameFirst", @"[А-Я][а-я]+\s[А-Я]\.\s?[А-Я]\."),
                new RegexPattern("InitialsFirst", @"[А-Я]\.\s?[А-Я]\.\s?[А-Я][а-я]+")
            );

            var codeExtractionStrategy = new OrganizationCodeRemovalStrategy();
            var nameRemovalStrategy = new RemoveReplacementStrategy();

            Console.WriteLine("Шаг 1: Стратегии поиска созданы");
            Console.WriteLine("Обозначения (3 паттерна)");
            Console.WriteLine("Имена (2 паттерна)\n");

            // ========================================================================
            // ШАГ 2: Конфигурация ПЕРВОГО ПРОХОДА
            // ========================================================================

            var firstPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>
                {
                    designationSearchStrategy,
                    nameSearchStrategy
                },
                ReplacementStrategy = new CompositeReplacementStrategy(
                    "FirstPassReplacement",
                    match => match.MatchType.Contains("Designation"),
                    codeExtractionStrategy,
                    nameRemovalStrategy
                ),
                Options = new ProcessingOptions
                {
                    ProcessProperties = true,
                    ProcessTextBoxes = true,
                    ProcessNotes = true,
                    ProcessHeaders = true,
                    ProcessFooters = true,
                    MinMatchLength = 5,
                    CaseSensitive = false
                },
                Logger = logger // ← ВАЖНО: Подключаем логгер
            };

            Console.WriteLine("Шаг 2: Конфигурация первого прохода");
            Console.WriteLine("Композитная стратегия замены");
            Console.WriteLine("Логгер подключен\n");

            // ========================================================================
            // ШАГ 3: Конфигурация ВТОРОГО ПРОХОДА
            // ========================================================================

            var secondPassConfig = new ProcessingConfiguration
            {
                SearchStrategies = new List<ITextSearchStrategy>(),
                ReplacementStrategy = new RemoveReplacementStrategy(),
                Options = new ProcessingOptions
                {
                    ProcessProperties = false,
                    ProcessTextBoxes = true,
                    ProcessNotes = true,
                    ProcessHeaders = true,
                    ProcessFooters = true,
                    MinMatchLength = 1,
                    CaseSensitive = true
                },
                Logger = logger
            };

            Console.WriteLine("Шаг 3: Конфигурация второго прохода\n");

            // ========================================================================
            // ШАГ 4: Двухпроходная конфигурация
            // ========================================================================

            var twoPassConfig = new TwoPassProcessingConfiguration
            {
                FirstPassConfiguration = firstPassConfig,
                SecondPassConfiguration = secondPassConfig,
                CodeExtractionStrategy = codeExtractionStrategy
            };

            Console.WriteLine("Шаг 4: Двухпроходная конфигурация создана\n");

            // ========================================================================
            // ШАГ 5: Запрос на обработку
            // ========================================================================

            var processingRequest = new DocumentProcessingRequest
            {
                InputFilePath = inputFile,
                OutputDirectory = outputDir,
                Configuration = firstPassConfig,
                ExportOptions = new ExportOptions
                {
                    ExportToPdf = false, // OpenXML не поддерживает PDF
                    SaveModified = true,
                    PdfFileName = null,
                    Quality = PdfQuality.HighQuality
                },
                PreserveOriginal = true
            };

            Console.WriteLine("Шаг 5: Запрос создан");
            Console.WriteLine($"Входной файл: {Path.GetFileName(inputFile)}");
            Console.WriteLine($"Выходная папка: {outputDir}");
            Console.WriteLine($"Сохранить оригинал: Да\n");

            // ========================================================================
            // ШАГ 6: ВЫПОЛНЕНИЕ С ИСПОЛЬЗОВАНИЕМ ResultAwareLogger
            // ========================================================================

            Console.WriteLine("Шаг 6: Начало обработки с ResultAwareLogger...\n");
            Console.WriteLine("=".PadRight(60, '='));

            using (var anonymizer = new DocumentAnonymizer(
                visible: false,
                useOpenXml: true, // Используем OpenXML
                logger: logger))
            {
                // Создаем ProcessingResult для автоматического сбора ошибок/предупреждений
                var result = new ProcessingResult();
                var resultAwareLogger = new ResultAwareLogger(logger, result);
                
                // Переподключаем логгеры в конфигурациях
                firstPassConfig.Logger = resultAwareLogger;
                secondPassConfig.Logger = resultAwareLogger;

                using (var factory = new DocumentProcessing.Documents.Factories.DocumentProcessorFactory(
                    visible: false,
                    useOpenXml: true,
                    logger: resultAwareLogger))
                {
                    var processor = factory.CreateProcessor(inputFile);

                    if (processor is ITwoPassDocumentProcessor twoPassProcessor)
                    {
                        result = twoPassProcessor.ProcessTwoPass(processingRequest, twoPassConfig);

                        // ========================================================================
                        // ШАГ 7: АНАЛИЗ РЕЗУЛЬТАТОВ
                        // ========================================================================

                        Console.WriteLine("=".PadRight(60, '='));
                        Console.WriteLine("\n--- РЕЗУЛЬТАТЫ ОБРАБОТКИ ---\n");

                        if (result.Success)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("✓ Обработка завершена УСПЕШНО!");
                            Console.ResetColor();
                            Console.WriteLine();

                            Console.WriteLine($"Статистика:");
                            Console.WriteLine($"Всего найдено совпадений: {result.MatchesFound}");
                            Console.WriteLine($"Успешно обработано: {result.MatchesProcessed}");

                            var extractedCodes = codeExtractionStrategy.GetExtractedCodes();

                            Console.WriteLine($"\nИзвлеченные коды организаций ({extractedCodes.Count}):");
                            foreach (var code in extractedCodes)
                            {
                                Console.WriteLine($"- {code}");
                            }

                            if (result.Metadata.TryGetValue("CodesRemoved", out var codesRemoved))
                            {
                                Console.WriteLine($"\nУдалено отдельно стоящих кодов: {codesRemoved}");
                            }

                            Console.WriteLine($"\nРезультаты сохранены в: {outputDir}");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("✗ Обработка завершена С ОШИБКАМИ!");
                            Console.ResetColor();
                            Console.WriteLine();

                            Console.WriteLine("Ошибки:");
                            foreach (var error in result.Errors)
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine($"{error}");
                                Console.ResetColor();
                            }
                        }

                        if (result.Warnings.Any())
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine($"\nПредупреждения ({result.Warnings.Count}):");
                            Console.ResetColor();
                            
                            foreach (var warning in result.Warnings)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.WriteLine($"{warning}");
                                Console.ResetColor();
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Процессор не поддерживает двухпроходную обработку");
                        result = processor.Process(processingRequest);

                        if (result.Success)
                        {
                            Console.WriteLine($"\nОбработано: {result.MatchesProcessed} совпадений");
                        }
                    }
                }
            }

            Console.WriteLine("\n=== ОБРАБОТКА ЗАВЕРШЕНА ===");
        }

        /// <summary>
        /// УПРОЩЕННЫЙ ПРИМЕР: Через фасад с логгированием
        /// </summary>
        public static void SimplifiedExample()
        {
            Console.WriteLine("=== УПРОЩЕННЫЙ ВАРИАНТ (через фасад) ===\n");

            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.SetMinimumLevel(LogLevel.Information);
            });

            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<DocumentAnonymizer>>();

            using (var anonymizer = new DocumentAnonymizer(
                visible: false,
                useOpenXml: true,
                logger: logger))
            {
                var result = anonymizer.AnonymizeDocumentWithCodeRemoval(
                    inputFilePath: @"/Users/paveldavydov/Documents/Test.docx",
                    outputDirectory: @"/Users/paveldavydov/Documents/Output"
                );

                if (result.Success)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"✓ Успех! Обработано: {result.MatchesProcessed}");
                    Console.ResetColor();

                    if (result.Metadata.TryGetValue("CodesRemoved", out var codesRemoved))
                    {
                        Console.WriteLine($"  Удалено кодов: {codesRemoved}");
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибки: {string.Join(", ", result.Errors)}");
                    Console.ResetColor();
                }
            }
        }

        /// <summary>
        /// ПРИМЕР: Пакетная обработка с подробным логгированием
        /// </summary>
        public static void BatchProcessingWithLogging()
        {
            Console.WriteLine("=== ПАКЕТНАЯ ОБРАБОТКА С ЛОГГИРОВАНИЕМ ===\n");

            var services = new ServiceCollection();
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddFilter("DocumentProcessing", LogLevel.Debug);
                builder.AddFilter("Microsoft", LogLevel.Warning);
            });

            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<DocumentAnonymizer>>();

            var files = new[]
            {
                @"/Users/paveldavydov/Documents/doc1.docx",
                @"/Users/paveldavydov/Documents/doc2.docx",
                @"/Users/paveldavydov/Documents/doc3.docx"
            };

            using (var anonymizer = new DocumentAnonymizer(logger: logger))
            {
                var config = DocumentAnonymizer.CreateDefaultConfiguration();
                config.Logger = logger;

                var batchResult = anonymizer.AnonymizeBatch(files, @"/Users/paveldavydov/Documents/Output", config);

                Console.WriteLine("\n=== ИТОГИ ПАКЕТНОЙ ОБРАБОТКИ ===");
                Console.WriteLine($"Всего файлов: {batchResult.TotalFiles}");
                Console.WriteLine($"Успешно: {batchResult.SuccessfulFiles}");
                Console.WriteLine($"Ошибок: {batchResult.FailedFiles}\n");

                foreach (var fileResult in batchResult.Results)
                {
                    if (fileResult.Success)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"✓ {fileResult.FileName}: {fileResult.MatchesProcessed} совпадений");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"{fileResult.FileName}: {fileResult.Error}");
                        Console.ResetColor();
                    }
                }
            }
        }
    }

    /// <summary>
    /// ТОЧКА ВХОДА
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = System.Text.Encoding.UTF8;

                // Выберите пример для запуска:

                // 1. Полная ручная настройка (РЕКОМЕНДУЕТСЯ)
                ManualConfigurationExample.FullManualConfiguration();

                // 2. Упрощенный вариант через фасад
                // ManualConfigurationExample.SimplifiedExample();

                // 3. Кастомные паттерны с Serilog
                // ManualConfigurationExample.CustomPatternsWithSerilog();

                // 4. Пакетная обработка
                // ManualConfigurationExample.BatchProcessingWithLogging();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\nКРИТИЧЕСКАЯ ОШИБКА: {ex.Message}");
                Console.WriteLine($"\nStack trace:\n{ex.StackTrace}");
                Console.ResetColor();
            }

            Console.WriteLine("\nНажмите любую клавишу для выхода...");
            Console.ReadKey();
        }
    }
}