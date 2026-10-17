"""Persian and English UI copy for WordCloner.

Wording is standardized for an administrative/document-issuance workflow.
"""


def _entry(persian, english):
    return {"persian": persian, "english": english}


class T:
    err_no_vars= _entry(
        (
        "در این قالب هیچ متغیری پیدا نشد.\n"
        "مطمئن شوید متغیرها با {w} احاطه شده‌اند "
        "(مثلاً {w}نام مشتری{w})."
    ),
    (
        "No variables were found in this template.\n"
        "Make sure variables are wrapped with {w} "
        "(for example {w}customer_name{w})."
    )
    )
    # Application and navigation
    app_title = _entry("سامانه صدور نامه", "Letter Issuance System")
    settings = _entry("تنظیمات", "Settings")
    settings_spaced = _entry("  تنظیمات", "  Settings")
    issue_letter = _entry("صدور نامه", "Issue Letter")
    add_new_letter = _entry("ثبت قالب نامهٔ جدید", "Register New Letter Template")
    back = _entry("بازگشت", "Back")
    back_spaced = _entry("  بازگشت", "  Back")

    # Template selection and quick access
    select_template = _entry("انتخاب قالب نامه", "Select a Letter Template")
    add_letter = _entry("ثبت قالب نامه", "Register Template")
    search_template = _entry("جست‌وجوی قالب نامه...", "Search letter templates...")
    untitled = _entry("بدون عنوان", "Untitled")
    drag_to_reorder = _entry(
        "برای تغییر ترتیب، بکشید و رها کنید.",
        "Drag and drop to change the order.",
    )

    quick_access = _entry("دسترسی سریع", "Quick Access")
    add_to_qa = _entry("مدیریت دسترسی سریع", "Manage Quick Access")
    add_template_to_qa = _entry(
        "افزودن قالب به دسترسی سریع", "Add Template to Quick Access"
    )
    remove_from_qa = _entry("حذف از دسترسی سریع", "Remove from Quick Access")
    no_qa_templates = _entry(
        "قالبی به دسترسی سریع افزوده نشده است؛ برای افزودن، «+» را انتخاب کنید.",
        "No templates have been added to Quick Access. Select “+” to add one.",
    )
    select_qa_hint = _entry(
        "قالب‌هایی را که باید در بخش «دسترسی سریع» نمایش داده شوند انتخاب کنید:",
        "Select the templates to display in Quick Access:",
    )
    no_templates_yet = _entry(
        "هنوز قالبی ثبت نشده است.", "No templates have been registered yet."
    )
    save = _entry("ذخیره", "Save")
    cancel = _entry("انصراف", "Cancel")

    # Word file selection
    drop_zone_hint = _entry(
        "فایل Word را اینجا رها کنید\nیا برای انتخاب از رایانه کلیک کنید",
        "Drop a Word file here\nor click to browse your computer",
    )
    select_word_file = _entry("انتخاب فایل Word", "Select a Word File")
    word_file_filter = _entry("سند Word (*.docx)", "Word Document (*.docx)")

    # Template fields and formatting
    font_size = _entry("اندازهٔ قلم", "Font Size")
    auto = _entry("خودکار", "Auto")
    bold = _entry("پررنگ", "Bold")
    italic = _entry("مورب", "Italic")
    color = _entry("رنگ متن", "Text Color")
    color_tooltip = _entry(
        "انتخاب رنگ متن؛ رنگ پیش‌فرض مشکی است.",
        "Choose a text color; the default is black.",
    )
    pick_color = _entry("انتخاب رنگ متن", "Select Text Color")
    font_size_tooltip = _entry(
        "اندازهٔ قلم را مشخص کنید. گزینهٔ «خودکار» اندازهٔ متن مجاور را اعمال می‌کند.",
        "Choose a font size. “Auto” matches the surrounding text.",
    )
    letter_name = _entry("نام قالب نامه", "Template Name")
    word_template_file = _entry("فایل Word قالب نامه", "Word Template File")
    output_filename = _entry("نام فایل خروجی", "Output File Name")
    variable_values = _entry("مقادیر متغیرها", "Variable Values")
    confirm_save = _entry("تأیید و ثبت قالب", "Confirm and Register Template")
    confirm_issue = _entry("تأیید و صدور نامه", "Confirm and Issue Letter")
    batch_excel_issue = _entry(
        "صدور جمعی از اکسل", "Bulk Issue from Excel"
    )
    batch_issue_title = _entry(
        "صدور جمعی از اکسل", "Bulk Issue from Excel"
    )
    batch_explanation = _entry(
        "در شیت فعال اکسل، ستون اول نام فایل Word خروجی است. "
        "ستون دوم به متغیر اول، ستون سوم به متغیر دوم و به همین ترتیب "
        "به همهٔ متغیرهای قالب اختصاص دارد. هر ردیف یک سند می‌سازد.\n"
        "سلول خالی برای متغیر خالی پذیرفته می‌شود. می‌توانید سطر اول را "
        "عنوان ستون‌ها قرار دهید و گزینهٔ مربوط را فعال کنید.",
        "In the active Excel sheet, column 1 is the output Word filename. "
        "Column 2 maps to the first template variable, column 3 to the second, "
        "and so on. Each row creates one document.\n"
        "An empty cell leaves that variable blank. You may use row 1 for column "
        "headings and enable the header-row option.",
    )
    batch_columns_title = _entry(
        "ترتیب ستون‌ها در اکسل", "Excel column mapping"
    )
    batch_format_title = _entry(
        "قالب‌بندی هر متغیر در سندهای خروجی", "Formatting for each variable"
    )
    batch_mapping_filename = _entry(
        "ستون {column}: نام فایل خروجی",
        "Column {column}: output filename",
    )
    batch_mapping_variable = _entry(
        "ستون {column}: متغیر «{name}»",
        "Column {column}: variable “{name}”",
    )
    batch_header_row = _entry(
        "سطر اول عنوان ستون‌هاست (در پردازش نادیده گرفته شود)",
        "Row 1 contains headings (skip it)",
    )
    batch_drop_hint = _entry(
        "فایل اکسل را اینجا رها کنید\nیا برای انتخاب از رایانه کلیک کنید",
        "Drop an Excel workbook here\nor click to browse your computer",
    )
    batch_excel_filter = _entry(
        "فایل اکسل (*.xlsx *.xlsm)", "Excel workbook (*.xlsx *.xlsm)"
    )
    batch_confirm = _entry(
        "تأیید و صدور جمعی", "Confirm and Create Documents"
    )
    batch_output_folder = _entry("صدور جمعی", "Bulk Issuance")
    batch_excel_missing = _entry(
        "ابتدا فایل اکسل را انتخاب یا در کادر مشخص‌شده رها کنید.",
        "Select or drop an Excel workbook first.",
    )
    batch_excel_library_missing = _entry(
        "کتابخانهٔ openpyxl نصب نشده است. برای نصب، دستور pip install openpyxl را اجرا کنید.",
        "The openpyxl library is missing. Install it with: pip install openpyxl",
    )
    batch_excel_not_found = _entry(
        "فایل اکسل یافت نشد:\n{path}",
        "Excel workbook not found:\n{path}",
    )
    batch_excel_type_error = _entry(
        "فایل باید با قالب .xlsx یا .xlsm باشد:\n{path}",
        "The workbook must be an .xlsx or .xlsm file:\n{path}",
    )
    batch_excel_read_error = _entry(
        "خواندن فایل اکسل ناموفق بود:\n{err}",
        "Could not read the Excel workbook:\n{err}",
    )
    batch_no_data_rows = _entry(
        "در شیت فعال، ردیف داده‌ای برای صدور پیدا نشد.",
        "No data rows were found in the active worksheet.",
    )
    batch_filename_empty = _entry(
        "نام فایل خروجی در سطر {row} خالی است.",
        "The output filename is empty on row {row}.",
    )
    batch_extra_columns = _entry(
        "سطر {row} در ستون {column} دادهٔ اضافه دارد. برای {variables} متغیر، "
        "پس از ستون نام فایل فقط {variables} ستون متغیر لازم است.",
        "Row {row} has extra data in column {column}. With {variables} variable(s), "
        "only {variables} variable column(s) may follow the filename column.",
    )
    batch_output_folder_error = _entry(
        "پوشهٔ خروجی صدور جمعی ایجاد یا قابل‌استفاده نیست:\n{err}",
        "Could not create or access the bulk output folder:\n{err}",
    )
    batch_row_write_error = _entry(
        "صدور سند در سطر {row} ناموفق بود:\n{err}",
        "Could not issue the document on row {row}:\n{err}",
    )
    batch_success = _entry(
        "صدور جمعی انجام شد. {count} سند در پوشهٔ زیر ذخیره شد:\n{path}",
        "Bulk issuance completed. {count} document(s) were saved in:\n{path}",
    )
    ph_letter_name = _entry(
        "برای مثال: نامهٔ پذیرش مقاله", "e.g., Article Acceptance Letter"
    )
    ph_filename = _entry(
        "برای مثال: نامه_پذیرش_مقاله", "e.g., Article_Acceptance_Letter"
    )
    no_variables = _entry(
        "برای این قالب متغیری تعریف نشده است.",
        "No variables are defined for this template.",
    )
    issue_named = _entry("صدور نامه: {name}", "Issue Letter: {name}")

    # Settings
    app_language = _entry("زبان رابط کاربری", "Interface Language")
    guide = _entry("  راهنما", "  Guide")
    guide_tooltip = _entry(
        "بازکردن فایل ویدیویی راهنمای WordCloner با پخش‌کنندهٔ پیش‌فرض.",
        "Open the local WordCloner video guide with the default player.",
    )
    launch_on_startup = _entry(
        "اجرای WordCloner هنگام ورود به Windows",
        "Launch WordCloner at Windows sign-in",
    )
    launch_on_startup_tooltip = _entry(
        "در زمان ورود به Windows، WordCloner در System Tray و بدون نمایش پنجره اجرا می‌شود. این گزینه به‌طور پیش‌فرض فعال است.",
        "At Windows sign-in, WordCloner starts in the system tray without showing its main window. This option is enabled by default.",
    )
    settings_contact_note = _entry(
        "چنانچه دربارهٔ استفاده از WordCloner پرسشی دارید، به راهنمایی نیاز دارید یا پیشنهادی برای بهبود برنامه در ذهن دارید، خوشحال می‌شویم از طریق ایمیل <a href=\"mailto:M.J.Nazify.Yummy@gmail.com\">M.J.Nazify.Yummy@gmail.com</a> با ما در میان بگذارید. از توجه و بازخورد ارزشمند شما سپاسگزاریم و با کمال میل پاسخ‌گوی شما خواهیم بود.",
        "If you have any questions about using WordCloner, need assistance, or have an idea or suggestion for improving the application, please feel free to email us at <a href=\"mailto:M.J.Nazify.Yummy@gmail.com\">M.J.Nazify.Yummy@gmail.com</a>. We truly appreciate your feedback and will be glad to hear from you.",
    )
    startup_update_error = _entry(
        "تنظیم اجرای خودکار WordCloner در Windows ناموفق بود:\n{err}",
        "Could not update the Windows startup setting:\n{err}",
    )
    output_path = _entry(
        "محل ذخیرهٔ نامه‌های صادرشده", "Output Folder for Issued Letters"
    )
    placeholder_wrapper = _entry("نشانگر متغیرها", "Placeholder Wrapper")
    browse = _entry("انتخاب...", "Browse...")
    save_settings = _entry("ذخیرهٔ تنظیمات", "Save Settings")
    select_output_dir = _entry("انتخاب پوشهٔ خروجی", "Select Output Folder")
    ph_wrapper = _entry("برای مثال: \"\"", 'e.g., ""')
    wrapper_hint = _entry(
        "متغیرها در فایل Word میان این نشانگر قرار می‌گیرند.\nنمونه: {example}",
        "In the Word file, place each variable between these wrappers.\nExample: {example}",
    )
    user_name = _entry("نام کاربر", "User Name")

    # Dialog buttons and messages
    ok = _entry("تأیید", "OK")
    yes = _entry("بله", "Yes")
    delete = _entry("حذف", "Delete")
    missing_library = _entry("کتابخانهٔ موردنیاز یافت نشد", "Required Library Not Found")
    no_docx = _entry(
        "کتابخانهٔ python-docx نصب نشده است.\n\nبرای نصب، دستور زیر را اجرا کنید:\npip install python-docx",
        "The python-docx library is not installed.\n\nInstall it by running:\npip install python-docx",
    )
    error = _entry("خطا", "Error")
    err_empty_name = _entry(
        "لطفاً نام قالب نامه را وارد کنید.", "Enter a name for the letter template."
    )
    err_empty_file = _entry(
        "لطفاً فایل Word قالب را انتخاب کنید یا آن را در ناحیهٔ مشخص‌شده رها کنید.",
        "Select a Word template file or drop it in the designated area.",
    )
    err_name_exists = _entry(
        "قالبی با نام «{name}» از پیش ثبت شده است.",
        "A template named “{name}” is already registered.",
    )
    saved = _entry("ثبت قالب انجام شد", "Template Registered")
    saved_with_n = _entry(
        "قالب نامهٔ «{name}» با {n} متغیر ثبت شد.",
        "Template “{name}” was registered with {n} variable(s).",
    )
    saved_no_vars = _entry(
        "یادآوری: در سند انتخاب‌شده، متغیری با نشانگر «{w}» یافت نشد.",
        "Note: No placeholders wrapped with “{w}” were found in the selected document.",
    )
    err_wrapper_count = _entry(
        "تعداد نشانگرهای متغیر با هم مطابقت ندارد",
        "Unmatched Placeholder Wrappers",
    )
    err_generic = _entry("خطایی رخ داد:\n{err}", "An error occurred:\n{err}")
    err_empty_filename = _entry(
        "لطفاً نام فایل خروجی را وارد کنید.", "Enter a name for the output file."
    )
    err_fill_one = _entry(
        "لطفاً مقدار دست‌کم یک متغیر را وارد کنید.",
        "Enter a value for at least one variable.",
    )
    err_output_dir_create = _entry(
        "ایجاد پوشهٔ خروجی انجام نشد:\n{err}",
        "Could not create the output folder:\n{err}",
    )
    err_template_not_found = _entry(
        "فایل قالب یافت نشد:\n{path}", "Template file not found:\n{path}"
    )
    file_exists = _entry("فایل از قبل وجود دارد", "File Already Exists")
    file_overwrite = _entry(
        "فایل «{f}» در پوشهٔ خروجی از قبل وجود دارد.\nآیا می‌خواهید آن را جایگزین کنید؟",
        "The file “{f}” already exists in the output folder.\nDo you want to replace it?",
    )
    issued = _entry("صدور نامه انجام شد", "Letter Issued")
    issued_ok = _entry(
        "نامه با موفقیت در مسیر زیر ذخیره شد:\n{path}",
        "The letter was issued successfully and saved to:\n{path}",
    )
    err_issue = _entry(
        "صدور نامه با خطا مواجه شد:\n{err}",
        "The letter could not be issued:\n{err}",
    )
    delete_template = _entry("حذف قالب نامه", "Delete Letter Template")
    confirm_delete_tpl = _entry(
        "آیا از حذف قالب نامهٔ «{name}» اطمینان دارید؟",
        "Are you sure you want to delete the template “{name}”?",
    )
    deleted = _entry("حذف قالب انجام شد", "Template Deleted")
    deleted_ok = _entry(
        "قالب نامهٔ «{name}» با موفقیت حذف شد.",
        "Template “{name}” was deleted successfully.",
    )
    err_output_empty = _entry(
        "لطفاً پوشهٔ خروجی را مشخص کنید.", "Select an output folder."
    )
    err_output_invalid = _entry(
        "پوشهٔ خروجی قابل استفاده نیست:\n{err}",
        "The output folder is not accessible:\n{err}",
    )
    err_wrapper_empty = _entry(
        "نشانگر متغیرها نمی‌تواند خالی باشد.",
        "The placeholder wrapper cannot be empty.",
    )
    var_count_error_detail = _entry(
        "تعداد نشانگرهای «{w}» در سند زوج نیست (تعداد: {n}).\n"
        "ممکن است نشانگر باز یا بستهٔ متناظر آن درج نشده باشد.\n"
        "لطفاً هر متغیر را به این صورت وارد کنید: {w}نام_متغیر{w}.",
        "The number of placeholder wrappers “{w}” in the document must be even "
        "(found: {n}).\nAn opening or closing wrapper may be missing.\n"
        "Write each variable as {w}variable_name{w}.",
    )

    # Command-line document issuance
    cli_values_required = _entry(
        "مقدارهای متغیرها را با گزینهٔ -v وارد کنید.",
        "Provide the variable values with -v.",
    )
    cli_too_many_values = _entry(
        "تعداد مقدارهای واردشده ({actual}) از تعداد متغیرهای قالب ({expected}) بیشتر است.",
        "Too many values: supplied {actual}, but the template has {expected} unique variable(s).",
    )
    cli_too_few_values = _entry(
        "تعداد مقدارهای واردشده ({actual}) از تعداد متغیرهای قالب ({expected}) کمتر است.\n"
        "برای متغیرِ خالی، جایگاه آن را حذف نکنید؛ به‌جای آن دقیقاً یک فاصله وارد کنید: \" \".",
        "Too few values: supplied {actual}, but the template has {expected} unique variable(s).\n"
        "Do not omit a blank variable's position; pass one space instead: \" \".",
    )
    cli_empty_value = _entry(
        "برای مقدار خالی از یک فاصلهٔ داخل نقل‌قول استفاده کنید (\" \")، نه رشتهٔ خالی (\"\").",
        "Use a quoted single space (\" \"), not an empty string (\"\"), for a blank value.",
    )
    cli_variable_count = _entry(
        "تعداد متغیرهای یکتای قالب: {count}",
        "Unique variables in template: {count}",
    )
    cli_count_mode_conflict = _entry(
        "گزینهٔ --count فقط با -i و -w قابل استفاده است؛ آن را با گزینه‌های صدور ترکیب نکنید.",
        "--count accepts -i and optional -w only; do not combine it with issuance options.",
    )
    cli_excel_mode_conflict = _entry(
        "برای صدور از اکسل، گزینه‌های -v و --force را وارد نکنید؛ مقدارها از فایل اکسل خوانده می‌شوند و نام‌های تکراری خودکار یکتا می‌شوند.",
        "Do not combine --excel with -v or --force; values come from the workbook and duplicate filenames are made unique.",
    )
    cli_output_required = _entry(
        "برای صدور تکی، مسیر فایل خروجی را با گزینهٔ -o مشخص کنید.",
        "For a single document, specify the output file with -o.",
    )
    cli_header_only_excel = _entry(
        "گزینهٔ --header-row فقط همراه با --excel قابل استفاده است.",
        "--header-row can only be used with --excel.",
    )
    cli_excel_output_folder_error = _entry(
        "در حالت اکسل، گزینهٔ -o باید مسیر یک پوشه باشد، نه فایل .docx:\n{path}",
        "In Excel mode, -o must be a folder path, not a .docx file:\n{path}",
    )
    cli_no_variables = _entry(
        "در فایل قالب هیچ متغیری پیدا نشد. مطمئن شوید متغیرها با نشانگر «{w}» نوشته شده‌اند.",
        "The template contains no variables. Check that each variable is wrapped with “{w}”.",
    )
    cli_value_count = _entry(
        "قالب {expected} متغیر یکتا دارد، اما {actual} مقدار وارد شده است.\n"
        "مقادیر را به ترتیب نخستین ظهور متغیرها در سند وارد کنید.",
        "The template has {expected} unique variable(s), but {actual} value(s) were supplied.\n"
        "Enter values in the order the variables first appear in the document.",
    )
    cli_input_not_file = _entry(
        "مسیر ورودی یک فایل نیست:\n{path}",
        "The input path is not a file:\n{path}",
    )
    cli_input_not_docx = _entry(
        "فایل ورودی باید با قالب .docx باشد:\n{path}",
        "The input must be a .docx Word document:\n{path}",
    )
    cli_output_is_input = _entry(
        "مسیر فایل خروجی با فایل قالب یکسان است؛ فایل قالب بازنویسی نمی‌شود.",
        "The output path is the same as the template; the template will not be overwritten.",
    )
    cli_output_is_directory = _entry(
        "مسیر خروجی به یک پوشه اشاره می‌کند، نه فایل:\n{path}",
        "The output path points to a directory, not a file:\n{path}",
    )
    cli_output_exists = _entry(
        "فایل خروجی از قبل وجود دارد:\n{path}\nبرای جایگزینی آن، گزینهٔ --force را اضافه کنید.",
        "The output file already exists:\n{path}\nUse --force to replace it.",
    )
    cli_output_directory_error = _entry(
        "پوشهٔ خروجی ایجاد یا قابل دسترسی نشد:\n{err}",
        "Could not create or access the output folder:\n{err}",
    )
    cli_read_error = _entry(
        "خواندن فایل قالب ناموفق بود:\n{err}",
        "Could not read the template:\n{err}",
    )
    cli_write_error = _entry(
        "ذخیرهٔ فایل خروجی ناموفق بود:\n{err}",
        "Could not save the output file:\n{err}",
    )
    cli_success = _entry(
        "سند با موفقیت در این مسیر ذخیره شد:\n{path}\nتعداد متغیرها: {count}",
        "Document created successfully:\n{path}\nVariables filled: {count}",
    )
    guide_not_found = _entry(
        "فایل ویدیوی راهنمای WordCloner یافت نشد:\n{path}",
        "WordCloner guide video was not found:\n{path}",
    )
    guide_open_error = _entry(
        "بازکردن فایل ویدیوی راهنما ناموفق بود:\n{err}",
        "Could not open the guide video:\n{err}",
    )
