import subprocess
import time

from contextlib import AbstractContextManager
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Optional, Tuple, TypeVar, Union, cast

import uno


__all__ = ["Cell", "Sheet", "Document"]

T = TypeVar("T")


def only_connected(function: Callable[..., T]) -> Callable[..., T]:
    """
    Helper decorator for Document class that
    only allows function call when Document is connected.
    """

    def wrapped(self: "Document", *args: Any, **kwargs: Any) -> T:
        if self.connected:
            return function(self, *args, **kwargs)
        raise ConnectionError('Call "connect" on this instance.')

    return wrapped


class BaseObject:
    """ Base class for all UNO objects. """

    def __init__(self, uno_obj: Any):
        self._uno_obj = uno_obj

    @property
    def name(self) -> str:
        return str(self._uno_obj.Name)

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.name})"

    def __str__(self) -> str:
        return self.name


class Cell(BaseObject):
    @property
    def value(self) -> str:
        return str(self._uno_obj.String)

    @value.setter
    def value(self, value: Any) -> None:
        if isinstance(value, (int, float)):
            self._uno_obj.Value = value
        else:
            self._uno_obj.String = str(value)

    @property
    def row_index(self) -> int:
        return int(self._uno_obj.RangeAddress.StartRow)

    @property
    def column_index(self) -> int:
        return int(self._uno_obj.RangeAddress.StartColumn)

    @property
    def column_name(self) -> str:
        return str(self._uno_obj.Columns.getByIndex(0).Name)

    @property
    def name(self) -> str:
        return f"{self.column_name}{self.row_index + 1}"

    @property
    def is_empty(self) -> bool:
        return str(self._uno_obj.Type.value) == "EMPTY"

    def parent(self) -> "Sheet":
        return Sheet(self._uno_obj.Spreadsheet)


class Sheet(BaseObject):
    def get_cell(self, cell_index: Union[str, Tuple[int, int]]) -> "Cell":
        try:
            if isinstance(cell_index, str):
                if ":" in cell_index:
                    raise Exception
                cell = self._uno_obj.getCellRangeByName(cell_index)
            elif isinstance(cell_index, tuple):
                cell = self._uno_obj.getCellByPosition(*cell_index)
            return Cell(cell)
        except:
            raise IndexError(f'"{cell_index}" not found')

    def __getitem__(self, key: Union[str, Tuple[int, int]]) -> "Cell":
        return self.get_cell(key)

    def append_row(self, values: Iterable[Any], offset: int = 0) -> None:
        rows = self._uno_obj.getRows()
        for row_index in range(rows.Count):
            if self[offset, row_index].is_empty:
                for col_index, value in enumerate(values):
                    cell = self[col_index + offset, row_index]
                    cell.value = value
                break

    def append_column(self, values: Iterable[Any], offset: int = 0) -> None:
        columns = self._uno_obj.getColumns()
        for col_index in range(columns.Count):
            if self[col_index, offset].is_empty:
                for row_index, value in enumerate(values):
                    cell = self[col_index, row_index + offset]
                    cell.value = value
                break


class Document(BaseObject, AbstractContextManager):
    def __init__(self, path: Union[str, Path], port: int = 2002) -> None:
        self.connected = False
        self._process: Optional[subprocess.Popen] = None
        self._path = path
        self._port = port

    def connect(self, max_tries: int = 10) -> None:
        self._process = subprocess.Popen(
            f'soffice --headless --accept="socket,host=localhost,port={self._port};'
            f'urp;StarOffice.ServiceManager" {self._path}',
            shell=True,
        )
        for _ in range(max_tries):
            try:
                local_context = uno.getComponentContext()
                resolver = local_context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", local_context
                )
                context = resolver.resolve(
                    f"uno:socket,host=localhost,port={self._port};urp;StarOffice.ComponentContext"
                )
                manager = context.ServiceManager
                desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
                self._uno_obj = desktop.getCurrentComponent()
                if self._uno_obj is not None:
                    self.connected = True
                    break
            except:
                pass
            finally:
                time.sleep(1)
        else:
            raise ConnectionError(f"Failed to connect to the document {self._path}")

    @only_connected
    def save(self) -> None:
        self._uno_obj.store()

    @only_connected
    def close(self) -> None:
        self._uno_obj.close(0)
        if self._process is not None:
            self._process.terminate()
            self._process.wait()

    @only_connected
    def get_sheet(self, sheet_id: Union[str, int]) -> "Sheet":
        """ Get sheet by index or name """
        try:
            if isinstance(sheet_id, int):
                return Sheet(self._uno_obj.Sheets.getByIndex(sheet_id))
            if isinstance(sheet_id, str):
                return Sheet(self._uno_obj.Sheets.getByName(sheet_id))
        except:
            raise IndexError(f'"{sheet_id}" not found')
        raise NotImplementedError(f"Key of type {type(sheet_id)} is not supported.")

    @only_connected
    def __iter__(self) -> Iterator["Sheet"]:
        iterator = self._uno_obj.Sheets.createEnumeration()
        while True:
            try:
                yield Sheet(iterator.nextElement())
            except:
                break

    def __getitem__(self, key: Union[str, int]) -> "Sheet":
        return self.get_sheet(key)

    def __enter__(self) -> "Document":
        self.connect()
        return self

    def __exit__(self, exc_type: Any, _exc_value: Any, _traceback: Any) -> None:
        if exc_type is None:
            self.save()
        self.close()

    @property
    def name(self) -> str:
        return str(self._uno_obj.Title)

    @property
    def sheet_names(self) -> Tuple[str, ...]:
        if not self.connected:
            return tuple()
        return cast(Tuple[str, ...], self._uno_obj.Sheets.ElementNames)

    def __len__(self) -> int:
        return len(self.sheet_names)
