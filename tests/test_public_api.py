from __future__ import annotations


def test_public_package_exports_supported_api():
    import skills.WPSComposer as api

    required = {
        "WriterComposer",
        "SheetComposer",
        "SlideComposer",
        "PdfComposer",
        "generate",
        "list_formats",
        "list_available_presets",
        "inspect",
        "edit",
        "attach_active",
        "parse",
        "parse_file",
        "StructuredDocument",
        "ConversionError",
        "convert_to_pdf",
    }

    assert required <= set(api.__all__)
    for name in required:
        assert hasattr(api, name), name


def test_public_import_does_not_require_pywin32():
    import skills.WPSComposer as api

    assert api.list_formats() == ["docx", "pptx", "xlsx", "pdf"]
    assert api.list_available_presets() == [
        "academic",
        "consultant",
        "business",
        "tech",
        "proposal",
    ]
