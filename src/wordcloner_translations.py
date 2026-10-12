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
