from pathlib import Path

p = Path(__file__).with_name("main.py")
text = p.read_text(encoding="utf-8")

marker = 'if __name__ == "__main__":'
idx = text.rfind(marker)
if idx < 0:
    raise SystemExit("marker not found")

new_block = '''if __name__ == "__main__":
    import faulthandler
    import traceback

    faulthandler.enable()

    def _excepthook(exc_type, exc_value, exc_tb):
        print("\\n========== UNCAUGHT EXCEPTION ==========", flush=True)
        traceback.print_exception(exc_type, exc_value, exc_tb)
        print("========================================\\n", flush=True)
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = _excepthook

    app = QApplication(sys.argv)
    window = EmailApp()
    window.show()
    sys.exit(app.exec())
'''

# Correct escaped newlines in print strings to real escape sequences for source
new_block = new_block.replace(
    'print("\\n========== UNCAUGHT EXCEPTION ==========", flush=True)',
    'print("\\n========== UNCAUGHT EXCEPTION ==========", flush=True)',
)

p.write_text(text[:idx] + new_block, encoding="utf-8")
print("patched ok")
print(p.read_text(encoding="utf-8")[-500:])
