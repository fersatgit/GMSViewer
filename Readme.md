<h1>GMSViewer</h1>
Утилита для просмотра и редактирования файлов c расширением GMS (CorelDraw global macro storage).
<p><img src=1.png><p>
<h2>Системные требования</h2>
Операционная система<td>Windows 2000 или выше
<h2>Как использовать</h2>
Ассоциируйте файлы с расширением gms с этой программой.
<h2>Что может</h2><ul>
<li>Отображать исходный код модулей, а также потоки PROJECT, VBFrame и немного dir, остальные потоки отображаются в hex-виде.
<li>Редактировать код модулей и несжатый текст (PROJECT и VBAFrame).
<li>Сохранять дамп потока в файл.
<li>Распаковывать сжатые данные из потока в файл.
<li>Удалять элементы дерева (не советую удалять что-либо кроме __SRP_XX потоков, при удалении модуля он не удаляется из dir).
<li>Перестравивать PROJECT. Это позволяет снять простую защиту в виде пароля или восстановить видимость скрытых модулей.
<li>Перестраивать dir поток, делая все модули доступными для редактирования. ВНИМАНИЕ - при перестройке будет удалена информация о справочных файлах (help context) и аргументы условной компиляции.
<li>Удалять ссылки на tlb и dll. В сочетании с опцией понижения версии это позволяет конвертировать простые gms из поздних версий CorelDraw в более ранние. Если gms использует типы данных из поздних версий, сконвертированный таким образом файл откроется в CorelDraw X3, но вы не сможете его скомпилировать.
<li>Удалять P-код. Удаляет P-код из модулей, очищет поток _VBA_PROJECT и удаляет _SRP_XX потоки. Это существенно уменьшает объём GMS, лечит инфицированные файлы.
<li>Понижать версию typeLib, заменяя содержимое потока "VBA Project Data". В сочетании с перестройкой dir потока позволяет открывать GMS в ранних версиях CorelDraw.</ul>
<h2>Пример понижения версии</h2>
Открываем файл <a href=https://forum.rudtp.ru/resources/makros-dlya-konvertirovaniya-iz-coreldraw-v-pdf-i-jpg.3226/download>Jpg_Pdf_Export.gms</a>, высталяем все галочки и нажимаем "сохранить". Теперь этот макрос может работать с CorelDraw X3. Конечно такое сработает не со всеми файлами. Макрос должен быть совместим на уровне исходного кода.
<p><img src=2.png><p>
<h2>Ссылки</h2><ol>
<li><a href=https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal>VBA Language Specification</a>
<li><a href=https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb>Compound File Binary File Format</a>
<li><a href=https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba>Office VBA File Format Structure</a>
<li><a href=https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms>Office Forms Binary File Formats</a>
<li><a href=https://github.com/bontchev/pcodedmp>Декомпилятор P-кода</a>. Для декомпиляции GMS нужно предварительно удалить 18 байт заголовка.
<li><a href=https://vbastomp.com>vbastomp</a> - разная информация по инфицированию VBA-файлов</ol>

