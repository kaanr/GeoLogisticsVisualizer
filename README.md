### GeoLogisticsVisualizer 🗺️
Интегрированный набор инструментов на Python для преобразования логистических данных из HTML в практические географические сведения. В этом репозитории содержатся скрипты для парсинга данных о местоположении из HTML-таблиц, их очистки и структурирования в DataFrame библиотеки pandas, а также экспорта в хорошо оформленные Excel-файлы. Также реализованы продвинутые возможности геокодирования для преобразования адресов в GPS-координаты и использования Folium для создания интерактивных карт, визуально представляющих маршруты доставки и места. Проект улучшает интерактивность карт за счет внедрения пользовательского JavaScript для динамичного взаимодействия с пользователем. Идеально подходит для аналитиков логистики, градостроителей и всех, кто занимается управлением и визуализацией географических данных в операциях доставки и перевозок.

### GeoLogisticsVisualizer 🗺️
An integrated Python toolkit for transforming HTML-based logistics data into actionable geographical insights. This repository contains scripts that parse location data from HTML tables, cleanse and structure it into a pandas DataFrame, and export it to a well-formatted Excel file. It also features advanced geocoding capabilities to convert addresses into GPS coordinates and uses Folium to create interactive maps that visually represent delivery routes and locations. Additionally, the project enhances map interactivity by injecting custom JavaScript for dynamic user interactions. Ideal for logistics analysts, urban planners, and anyone involved in managing and visualizing geographical data for delivery and shipping operations.

#### Немного предыстории:
В один весенний день, столкнувшись с необходимостью найти подработку, я решил использовать свою машину для доставки товаров с маркетплейсов. Все началось довольно банально: мне нужно было дополнительное финансирование, и такая работа показалась идеальной возможностью совместить полезное с приятным — зарабатывать, наслаждаясь вечерними поездками по городу.

Первый день работы выдался особенно тяжёлым. Работодатель предоставил мне пять листов с таблицами, напечатанными в таком хаотичном порядке, что сориентироваться в них было почти невозможно. Адреса были перемешаны, информация о посылках — разбросана по всему документу. Check my [test document](./GeoLogisticsVisualizer/Маршрутный лист.html)!Мгновенно стало ясно, что без порядка мне не удастся эффективно планировать маршруты и выполнять работу в срок.

В этот момент я решил взять инициативу в свои руки. Используя свои навыки программирования, я написал скрипт, который автоматизировал бы обработку этих данных. Программа читала информацию из HTML-файлов, преобразовывала их в удобные таблицы, сортировала по нужным критериям и, что самое важное, определяла географические координаты каждого адреса.

Так я не только значительно ускорил процесс подготовки к каждой поездке, но и смог предложить своему работодателю решение, которое в дальнейшем было принято на вооружение всеми курьерами компании. Этот опыт не только помог мне лучше устроить свою вечернюю подработку, но и добавил ценный проект в мой портфолио.
