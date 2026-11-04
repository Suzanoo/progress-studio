"""P6 Cost UDF byte patches; never serialize a normalized schedule.

Expat supplies namespace-aware identity and byte offsets. Untouched bytes,
including comments, processing instructions and foreign extensions, survive.
Only UTF-8 APIBusinessObjects documents without DTDs are supported in XW-1.
"""
from dataclasses import dataclass, field
import re
from xml.parsers import expat
from xml.sax.saxutils import escape

from progress_studio.domain.amount_field import AmountField
from progress_studio.domain.xml_writeback import monetary_decimal

TITLE = 'Progress Studio Amount'
XSI_NIL = 'http://www.w3.org/2001/XMLSchema-instance}nil'


@dataclass
class _Node:
    tag: str
    attrs: dict
    start: int
    open_end: int
    qname: str
    closed: bool
    end_start: int = 0
    end: int = 0
    text: str = ''
    children: list = field(default_factory=list)


@dataclass(frozen=True)
class Patch:
    start: int
    end: int
    replacement: bytes


def _parse(data):
    data.decode('utf-8-sig')  # Do not transcode original bytes.
    parser = expat.ParserCreate(namespace_separator='}')
    stack, nodes = [], []

    def forbidden(*args):
        raise ValueError('DTD/entity declarations are unsupported for writeback.')

    def declaration(version, encoding, standalone):
        if encoding and encoding.lower().replace('-', '') != 'utf8':
            raise ValueError('Writeback requires UTF-8 XML.')

    def start(tag, attrs):
        offset = parser.CurrentByteIndex
        # Scan the tag respecting quotes (an attribute may contain >).
        quote = None
        pos = offset
        while pos < len(data):
            char = data[pos]
            if quote:
                if char == quote:
                    quote = None
            elif char in (34, 39):
                quote = char
            elif char == 62:
                break
            pos += 1
        token = data[offset:pos + 1]
        qname = re.match(rb'<([^\s/>]+)', token).group(1).decode('utf-8')
        node = _Node(tag, attrs, offset, pos + 1, qname, token.rstrip().endswith(b'/>'))
        if stack:
            stack[-1].children.append(node)
        nodes.append(node)
        stack.append(node)

    def end(tag):
        node = stack.pop()
        node.end_start = node.open_end if node.closed else parser.CurrentByteIndex
        node.end = node.open_end if node.closed else data.index(b'>', node.end_start) + 1

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = lambda text: setattr(stack[-1], 'text', stack[-1].text + text) if stack else None
    parser.StartDoctypeDeclHandler = forbidden
    parser.EntityDeclHandler = forbidden
    parser.ExternalEntityRefHandler = forbidden
    parser.XmlDeclHandler = declaration
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:
        raise ValueError(f'Invalid P6 XML: {exc}') from exc
    return nodes[0], nodes


class P6AmountDocument:
    def __init__(self, data: bytes):
        self.data = data
        self.root, self.nodes = _parse(data)
        self.namespace, _, local = self.root.tag.rpartition('}')
        if local != 'APIBusinessObjects' or not re.fullmatch(
            r'http://xmlns\.oracle\.com/Primavera/(?:P6Professional|P6)/[^/]+/API/BusinessObjects', self.namespace
        ):
            raise ValueError('Expected namespaced Primavera P6 APIBusinessObjects XML.')
        projects = [n for n in self.nodes if n.tag == self.tag('Project')]
        if len(projects) != 1 or projects[0] not in self.root.children:
            raise ValueError('Exactly one direct Project is required; multi-Project targets are unsupported.')
        self.project = projects[0]
        self.activities = {}
        seen = set()
        objects = set()
        for node in self.children(self.project, 'Activity'):
            aid = self.text(node, 'Id')
            key = aid.casefold()
            if not aid or key in seen:
                raise ValueError('Blank/duplicate/case-colliding target Activity ID: ' + aid)
            seen.add(key)
            oid = self.text(node, 'ObjectId')
            if oid and oid in objects:
                raise ValueError('Duplicate Activity ObjectId: ' + oid)
            if oid:
                objects.add(oid)
            self.activities[aid] = node
        self.definitions = [n for n in self.nodes if n.tag == self.tag('UDFType')]
        definition_ids = set()
        for node in self.definitions:
            oid = self.text(node, 'ObjectId')
            if not re.fullmatch(r'[1-9][0-9]*', oid) or oid in definition_ids:
                raise ValueError('Missing, noncanonical or duplicate UDF declaration ObjectId: ' + oid)
            definition_ids.add(oid)

    def tag(self, name):
        return self.namespace + '}' + name

    def children(self, node, name):
        return [n for n in node.children if n.tag == self.tag(name)]

    def one(self, node, name):
        matches = self.children(node, name)
        if len(matches) > 1:
            raise ValueError('Ambiguous duplicate ' + name)
        return matches[0] if matches else None

    def text(self, node, name):
        child = self.one(node, name)
        if child is None:
            return ''
        if child.children:
            raise ValueError('Non-scalar ' + name)
        return child.text.strip()

    def fields(self):
        return tuple(AmountField('p6:udf:' + self.text(n, 'ObjectId'), self.text(n, 'Title'),
                                 'P6', self.text(n, 'DataType')) for n in self.definitions
                     if self.text(n, 'SubjectArea') == 'Activity')

    def resolve_field(self, identity: str | None, create: bool):
        if create:
            if identity:
                raise ValueError('Choose an existing field OR request creation, not both.')
            if any(self.text(n, 'Title').casefold() == TITLE.casefold() for n in self.definitions):
                raise ValueError('Progress Studio Amount already exists; explicitly select a compatible identity.')
            # Include dangling references and other object IDs: no example constant.
            used = {n.text.strip() for n in self.nodes if n.tag in
                    (self.tag('ObjectId'), self.tag('TypeObjectId'))}
            number = max((int(s) for s in used if s.isdecimal()), default=0) + 1
            if number > 2147483647:
                raise ValueError('No safe supported new UDF ObjectId available.')
            return AmountField(f'p6:udf:{number}', TITLE, 'P6', 'Cost'), True
        if not identity or not re.fullmatch(r'p6:udf:[1-9][0-9]*', identity):
            raise ValueError('Explicitly select a typed P6 Activity Cost UDF.')
        oid = identity.split(':')[-1]
        matches = [n for n in self.definitions if self.text(n, 'ObjectId') == oid]
        if len(matches) != 1:
            raise ValueError('Missing or duplicate selected UDF declaration: ' + identity)
        node = matches[0]
        if node not in self.root.children or self.text(node, 'SubjectArea') != 'Activity' or self.text(node, 'DataType') != 'Cost':
            raise ValueError('Selected field must be a direct, declared Activity Cost UDF.')
        return AmountField(identity, self.text(node, 'Title'), 'P6', 'Cost'), False

    def milestone(self, aid):
        kind = self.text(self.activities[aid], 'Type')
        if kind in ('Start Milestone', 'Finish Milestone'):
            return True
        if kind in ('Task Dependent', 'Resource Dependent', 'Level of Effort', 'WBS Summary'):
            return False
        raise ValueError(f'{aid}: missing/unsupported target Activity Type {kind!r}.')

    def value_node(self, aid, identity):
        oid = identity.split(':')[-1]
        values = self.children(self.activities[aid], 'UDF')
        for node in values:
            ref = self.text(node, 'TypeObjectId')
            if not re.fullmatch(r'[1-9][0-9]*', ref):
                raise ValueError(f'{aid}: missing/noncanonical UDF TypeObjectId.')
        matches = [n for n in values if self.text(n, 'TypeObjectId') == oid]
        if len(matches) > 1:
            raise ValueError(f'{aid}: duplicate selected UDF values.')
        udf = matches[0] if matches else None
        if udf is None:
            return None, None, None
        if any(n.tag != self.tag('CostValue') and n.tag.rsplit('}', 1)[-1].endswith('Value') for n in udf.children):
            raise ValueError(f'{aid}: incompatible typed UDF value.')
        cost = self.one(udf, 'CostValue')
        if cost is None:
            return udf, None, None
        if cost.children or any(key != XSI_NIL for key in cost.attrs):
            raise ValueError(f'{aid}: unsupported CostValue content/attributes.')
        nil = cost.attrs.get(XSI_NIL, 'false')
        if nil not in ('false', '0', 'true', '1') or (nil in ('true', '1') and cost.text.strip()):
            raise ValueError(f'{aid}: ambiguous nil CostValue.')
        raw = None if nil in ('true', '1') else cost.text.strip() or None
        return udf, cost, raw

    @staticmethod
    def fragment(parent, name, text=None, inner=None):
        prefix = parent.qname.rsplit(':', 1)[0] + ':' if ':' in parent.qname else ''
        qname = prefix + name
        body = escape(text) if text is not None else inner or ''
        return f'<{qname}>{body}</{qname}>'

    def _insert(self, parent, fragment):
        if parent.closed:
            # Replace only the empty closing delimiter; preserve all attributes.
            return Patch(parent.open_end - 2, parent.open_end,
                         ('>' + fragment + '</' + parent.qname + '>').encode())
        return Patch(parent.end_start, parent.end_start, fragment.encode())

    def patches(self, changes, selected, create):
        patches = []
        oid = selected.identity.split(':')[-1]
        if create:
            body = ''.join(self.fragment(self.root, k, v) for k, v in (
                ('DataType', 'Cost'), ('IsSecureCode', '0'), ('ObjectId', oid),
                ('SubjectArea', 'Activity'), ('Title', TITLE)))
            declaration = self.fragment(self.root, 'UDFType', inner=body).encode()
            existing = self.children(self.root, 'UDFType')
            currency = self.children(self.root, 'Currency')
            at = (existing or currency)[-1].end if existing or currency else self.root.open_end
            patches.append(Patch(at, at, declaration))
        for change in changes:
            if change.status != 'replace':
                continue
            activity = self.activities[change.activity_id]
            udf, cost, _ = self.value_node(change.activity_id, selected.identity)
            if cost is not None:
                if not cost.closed and not cost.attrs:
                    patches.append(Patch(cost.open_end, cost.end_start, change.after_raw.encode()))
                else:
                    opening = self.data[cost.start:cost.open_end]
                    # Retain local namespace declarations on the value element.
                    # Only xsi:nil belongs to the changed numeric representation.
                    if XSI_NIL in cost.attrs:
                        opening = re.sub(rb'\s+[^\s=<>]+:nil\s*=\s*(?:"[^"]*"|\x27[^\x27]*\x27)', b'', opening)
                    opening = re.sub(rb'/\s*>$', b'>', opening)
                    replacement = opening + change.after_raw.encode() + ('</' + cost.qname + '>').encode()
                    patches.append(Patch(cost.start, cost.end, replacement))
            elif udf is not None:
                patches.append(self._insert(udf, self.fragment(udf, 'CostValue', change.after_raw)))
            else:
                body = self.fragment(activity, 'TypeObjectId', oid) + self.fragment(activity, 'CostValue', change.after_raw)
                patches.append(self._insert(activity, self.fragment(activity, 'UDF', inner=body)))
        return tuple(sorted(patches, key=lambda p: (p.start, p.end)))

    def render(self, changes, selected, create):
        patches = self.patches(changes, selected, create)
        chunks, position = [], 0
        for patch in patches:
            if patch.start < position:
                raise ValueError('Overlapping XML updates.')
            chunks.extend((self.data[position:patch.start], patch.replacement))
            position = patch.end
        chunks.append(self.data[position:])
        output = b''.join(chunks)
        # Independently parse and resolve the emitted target/value structure.
        check = P6AmountDocument(output)
        check.resolve_field(selected.identity, False)
        for change in changes:
            _, _, raw = check.value_node(change.activity_id, selected.identity)
            if change.status == 'preserve_milestone':
                if raw != change.before_raw:
                    raise ValueError('Milestone preservation verification failed.')
            elif raw is None or monetary_decimal(raw) != change.source_amount:
                raise ValueError('Written Amount verification failed: ' + change.activity_id)
        return output
