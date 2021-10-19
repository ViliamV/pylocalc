import subprocess
import time

from contextlib import AbstractContextManager
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Iterable, Iterator, Optional, Tuple, TypeVar, Union, cast

import uno


T = TypeVar("T")


def _only_connected(function: Callable[..., T]) -> Callable[..., T]:
    """
    Helper decorator for Document class that
    only allows function call when Document is connected.
    """

    def wrapped(document: "Document", *args: Any, **kwargs: Any) -> T:
        if document.connected:
            return function(document, *args, **kwargs)
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
        """ Value of the cell. """
        return str(self._uno_obj.String)

    @value.setter
    def value(self, value: Any) -> None:
        if isinstance(value, (int, float)):
            self._uno_obj.Value = value
        elif isinstance(value, Decimal):
            self._uno_obj.Value = float(value)
        else:
            self._uno_obj.String = str(value)

    @property
    def row_index(self) -> int:
        """ Zero-based row index of the cell. """
        return int(self._uno_obj.RangeAddress.StartRow)

    @property
    def column_index(self) -> int:
        """ Zero-based column index of the cell. """
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
        empty_index = self.find_index("", column=offset, find_empty=True)
        for col_index, value in enumerate(values):
            cell = self[col_index + offset, empty_index]
            cell.value = value

    def append_column(self, values: Iterable[Any], offset: int = 0) -> None:
        empty_index = self.find_index("", row=offset, find_empty=True)
        for row_index, value in enumerate(values):
            cell = self[empty_index, row_index + offset]
            cell.value = value

    def find_index(
        self,
        value: str,
        *,
        row: Optional[int] = None,
        column: Optional[int] = None,
        start: Optional[int] = None,
        end: Optional[int] = None,
        find_empty: bool = False,
    ) -> int:
        """
        Return zero-based index of the first cell whose value is equal to `value` in either `row` or `column`.
        Raises a ValueError if there is no such cell.
        The returned index is computed relative to the beginning of the row or column rather than the `start` argument.
        :param value: value to search for
        :param row: index of row to search in if not None
        :param column: index of column to search in if not None
        :param start: restrict search to start from this index (inclusive)
        :param end: restrict search to end at this index (exclusive)
        :param find_empty: ignore `value` and search for empty cell
        """
        if row is None and column is None:
            raise IndexError("Both `row` and `column` cannot be `None`.")
        if row is not None and column is not None:
            raise IndexError("Both `row` and `column` cannot have values.")
        matches = lambda cell: (cell.is_empty() if find_empty else cell.value == value)
        fixed_index = row or column
        count = int(self._uno_obj.Columns.Count) if row is not None else int(self._uno_obj.Rows.Count)
        min_index = max(0, start or 0)
        max_index = min(count, end or count)
        for index in range(min_index, max_index):
            if (row is not None and self[index, row].value == value) or (
                column is not None and self[column, index].value == value
            ):
                return index
        raise ValueError("Not found.")


class Document(BaseObject, AbstractContextManager):
    def __init__(self, path: Union[str, Path], port: int = 2002, host: str = "localhost") -> None:
        self.connected = False
        self._process: Optional[subprocess.Popen] = None
        self._path = path
        self._port = port
        self._host = host
        super().__init__(None)

    def connect(self, max_tries: int = 10) -> None:
        self._process = subprocess.Popen(
            f'soffice --headless --accept="socket,host={self._host},port={self._port};'
            f'urp;StarOffice.ServiceManager" "{self._path}"',
            shell=True,
        )
        last_error = None
        for _ in range(max_tries):
            try:
                local_context = uno.getComponentContext()
                resolver = local_context.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", local_context
                )
                context = resolver.resolve(
                    f"uno:socket,host={self._host},port={self._port};urp;StarOffice.ComponentContext"
                )
                manager = context.ServiceManager
                desktop = manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
                self._uno_obj = desktop.getCurrentComponent()
                if self._uno_obj is not None:
                    self.connected = True
                    break
            except Exception as e:
                last_error = e
                time.sleep(1)
        else:
            raise ConnectionError(
                f"Failed to connect to the document {self._path}\n"
                "Try to run `ps ax | grep 'soffice --headless'` and kill the running process."
            ) from last_error

    @_only_connected
    def save(self) -> None:
        self._uno_obj.store()

    @_only_connected
    def close(self) -> None:
        self._uno_obj.close(0)
        if self._process is not None:
            self._process.terminate()
            self._process.wait()

    @_only_connected
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

    @_only_connected
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
        if self._uno_obj is not None:
            return str(self._uno_obj.Title)
        return str(self._path)

    @property
    def sheet_names(self) -> Tuple[str, ...]:
        if not self.connected:
            return tuple()
        return cast(Tuple[str, ...], self._uno_obj.Sheets.ElementNames)

    def __len__(self) -> int:
        return len(self.sheet_names)
