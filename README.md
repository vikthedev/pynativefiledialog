# pynativefiledialog

Native Windows file dialogs (Vista+) for Python using COM.

`pynativefiledialog` is a lightweight, dependency-free wrapper around the
native Windows File Open / Save dialogs. It provides direct access to the
Explorer-style dialogs through the Windows COM API.

The library is implemented as a single module for easy embedding and flexible
import usage.

---

## Features

- Native Windows dialogs (Explorer-style)
- Windows Vista or newer (Windows 6.0+)
- File open and save dialogs
- Folder selection
- Multi-file selection
- File type filters
- Custom titles and button labels
- No external dependencies

---

## Requirements

- Windows Vista or newer
- Python 3.9+

---

## Installation

```bash
pip install pynativefiledialog
```

---

## Basic Usage

### Open a single file

```python
from pynativefiledialog import NativeFileDialog

path = NativeFileDialog.get_file(title="Select a file")
```

### Open multiple files

```python
from pynativefiledialog import NativeFileDialog

paths = NativeFileDialog.get_files(title="Select files")
```

### Save file dialog

```python
from pynativefiledialog import NativeFileDialog

path = NativeFileDialog.set_file(title="Save output file")
```

### Select a folder

```python
from pynativefiledialog import NativeFileDialog

folder = NativeFileDialog.get_dir(title="Select folder")
```

---

## File Filters

### Using predefined filters

```python
from pynativefiledialog import CommonFilters, NativeFileDialog

path = NativeFileDialog.get_file(
    title="Select image",
    file_type_filters=(
        CommonFilters.IMAGE_ALL.filter,
        CommonFilters.PNG.filter,
        CommonFilters.ALL.filter,
    )
)
```

### Custom filters

```python
from pynativefiledialog import FileFilter, NativeFileDialog

filters = (
    FileFilter("Images", ("png", "jpg", "jpeg")),
    FileFilter("All files", ("*.*",)),
)

path = NativeFileDialog.get_file(file_type_filters=filters)
```

---

## License

MIT License
