from xls_management.tui.project_form import ProjectChoice


def project_combo_box() -> tuple[str, bool]:
    app = ProjectChoice()
    app.run()
    return app.project_name, app.evalue_master_id

if __name__ == "__main__":
    name, evalue = ProjectChoice().run()
    print(f"project: {name}\nevalue id master: {evalue}")