# git-tasks-parser

CLI-утилита создаёт Excel-отчёт по всем коммитам, достижимым из выбранной ветки Git. Обычные слияния включены; коммиты из ещё не слитых веток исключены.

```powershell
uv sync
uv run git-tasks-parser --repo C:\work\project --main-branch main --output C:\reports\commits.xlsx
```

Если зависимости уже установлены, программу можно запустить без сборки пакета: `uv run --no-sync python main.py --repo C:\work\project --main-branch main --output C:\reports\commits.xlsx`.

Все три параметра обязательны. Короткое имя ветки означает локальную ветку (`refs/heads/main`); для remote-tracking ветки укажите полное имя, например `refs/remotes/origin/main`. Утилита использует локальное состояние Git и не делает `fetch`. Существующий файл результата не перезаписывается.

В отчёте один лист с колонками `Commit hash`, `Issue key`, `author email`. Для каждого уникального ключа задачи в полном сообщении коммита создаётся отдельная строка. Коммит без ключа тоже получает строку с пустым полем `Issue key`. Ключи распознаются в верхнем регистре по шаблону вроде `PROJ-123` или `MAIN_2-4`. Значения ячеек записываются как текст.

Проверка:

```powershell
uv run --no-sync python -m unittest discover -s tests
```
