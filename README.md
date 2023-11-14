# Report Box

Report Box – это консольный, хардкорный, но очень эффективный Python-скрипт, позволяющий делать в основном три вещи: 1) очень быстро вводить отчеты возвещателей с помощью автоматизации нажатий на клавиатуру; 2) смотреть, кто еще не сдал отчет; 3) генерировать общий отчет для ввода на сайте. Программа подходит тем, кто дружит с клавиатурой и способен потратить 5 минут на чтение документации. В принципе, эти качества должны быть у каждого секретаря. 😅

![image](https://github.com/antorix/Report-Box/assets/9825468/7f398a48-8025-4d78-8e70-5d03aec78f9b)

## Установка и начало работы

1. Если в системе нет Python, установите его. Прямая ссылка для Windows: [python-3.8.6-amd64-webinstall.exe](https://www.python.org/ftp/python/3.8.6/python-3.8.6-amd64-webinstall.exe). (На Linux программа работает, но необходимо запускать в режиме суперпользователя.)
2. Скачайте и запустите файл программы: [reportbox.py](https://github.com/antorix/Report-Box/releases/download/release/reportbox.py).
3. Программа создаст папки `Возвещатели`, `Подсобные пионеры` и `Общие пионеры`. Скопируйте в них PDF-файлы соответствующих возвещателей.
4. В качестве просмотрщика PDF-файлов по умолчанию рекомендуется Microsoft Edge или [PDF-XChange Editor](https://www.pdf-xchange.de/DL/tracker10/editor-msi64-tracker.php). (Можете пробовать и другие просмотрщики, но в них может некорректно работать или вообще не работать автоматизация вставки данных. В двух вышеупомянутых программах она протестирована и работает идеально.)

Можно начинать работать! Вверху экрана есть подсказки по всем командам. В принципе, этого достаточно, но еще несколько пояснений, если вам интересно углубиться в детали или не все понятно.

## Принципы работы

Идею Report Box можно описать так: пульт управления PDF-бланками S-21. Вы из одного места осуществляете все операции с картотекой бланков: создание, удаление, перемещение, переименовывание, ввод отчетов и анализ их статистики. С PDF-файлами напрямую вы больше не работаете.

Однако важно понимать, что программа использует уже существующие официальные бланки и не пытается генерировать их сама. Это безопасно и не создает зависимость от отдельно взятой программы. Даже база данных находится в обычном текстовом файле. Rocket Box – это своего рода надстройка над картотекой бланков, но не их замена.

> **Внимание! На данный момент Report Box работает с бланками S-21 старого образца. Версия приложения под бланк нового образца (ноябрь 2023) выйдет 30 ноября.**

Программа создает и поддерживает собственную базу данных путем сканирования своих папок с бланками, и все операции выполняются с этой базой. Она находится в файле `publishers.ini`. Содержимое ваших PDF-файлов программа не видит и не изменяет (кроме единственного случая ввода отчета, описанного ниже). Если скопировать в эти папки новые бланки, они сразу появятся в программе (программа постоянно мониторит папки).

Все дальнейшие манипуляции с бланками делаются только в программе, настоятельно не рекомендуется выполнять их внутри папок вручную, это может привести к неожиданным для вас последствиям. Для создания новых возвещателей нужно указать в программе местоположение чистого бланка `S-21_U.pdf`. Если удалить файл `publishers.ini`, программа сгенерирует базу заново – так можно начать с чистого листа.

### Как открыть возвещателя

Чтобы открыть любого возвещателя, вы просто вводите название его файла (без части `.PDF`) и жмете Enter. Можно ввести даже частично. Скажем, если файл называется `1КА.pdf`, то можно ввести и `КА`, и даже `А` (регистр не важен). Если будет несколько совпадений, вам будет предложено выбрать из списка. Затем появится меню действий, которые можно выполнить с возвещателем. Выбор осуществляется путем ввода номера и нажатия Enter.

### Как ввести отчет

Это самая важная часть. Отчет вводится в одной-единственной строке. Сначала вводите название файла, как упоминалось выше, а затем через пробел – количество часов. Например: `Иван 5`. Если нужно ввести изучение, вводите его дальше через еще один пробел: `Иван 5 1`. Если это пионер и у него есть кредит часов, вводите его следующим: `Иван 50 1 10`. Если любое значение отсутствует, вводите ноль. (Вместо пробелов можно использовать табуляцию.)

Когда вы нажимаете `Enter`, программа сама открывает PDF-файл этого возвещателя. А затем происходит главное волшебство: **поставьте курсор мыши в поле «Часы» соответствующего месяца и нажмите на кнопку «Insert» на клавиатуре**. Вы увидите, как отчет целиком ввелся в PDF-файл. Файл также сам сохраняется. Вам остается только закрыть его или дописать примечание, если нужно.

Если у вас остались вопросы или что-то не работает, пишите на [inoblogger@gmail.com](mailto:inoblogger@gmail.com).
