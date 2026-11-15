"""Desktop selection state. Financial values and visibility belong to PR-1."""
import json
import os
from pathlib import Path
from tempfile import NamedTemporaryFile

from progress_studio.presentation.gui.strings import tr

SUFFIX = '.paymentselection.json'


class PlanSelection:
    def __init__(self, source):
        self.source = source
        self.by_key = {row.key: row for row in source.rows}
        self.keys = frozenset()
        self.result = source.select((), revision=source.revision)

    def select(self, keys):
        wanted = frozenset(keys)
        if wanted != self.keys:
            result = self.source.select(wanted, revision=self.source.revision)
            self.keys, self.result = wanted, result

    def matching(self, query):
        text = query.strip().casefold()
        if not text:
            return self.source.rows
        visible = set()
        for row in self.source.rows:
            if any(text in field.casefold() for field in (row.code, row.name, row.key)):
                key = row.key
                while key is not None and key not in visible:
                    visible.add(key)
                    key = self.by_key[key].parent
        return tuple(row for row in self.source.rows if row.key in visible)

    def save(self, path):
        path = Path(path)
        if not path.name.lower().endswith(SUFFIX):
            raise ValueError(tr('pr2.filetype'))
        data = {'format': 'progress-studio.payment-selection', 'version': 1,
                'project': self.source.project, 'revision': self.source.revision,
                'workbook': self.source.filename, 'rows': [r.key for r in self.result.rows]}
        temp = None
        try:
            with NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent,
                                    delete=False) as stream:
                temp = Path(stream.name)
                json.dump(data, stream, ensure_ascii=False, indent=2)
            os.replace(temp, path)
        finally:
            if temp:
                temp.unlink(missing_ok=True)

    def restore(self, path):
        data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
        if (not isinstance(data, dict) or data.get('format') != 'progress-studio.payment-selection'
                or type(data.get('version')) is not int or data['version'] != 1):
            raise ValueError(tr('pr2.invalid'))
        if (data.get('project'), data.get('revision')) != (self.source.project, self.source.revision):
            raise ValueError(tr('pr2.revision'))
        keys = data.get('rows')
        if (not isinstance(keys, list) or any(not isinstance(k,str) or k not in self.by_key for k in keys)
                or len(keys) != len(set(keys))):
            raise ValueError(tr('pr2.invalid'))
        self.select(keys)
