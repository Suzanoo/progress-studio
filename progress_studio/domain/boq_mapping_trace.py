"""Read-only relationship indexes over the accepted saved-share snapshot."""
from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class MappingRelationship:
    boq_key: str
    boq_description: str
    activity_id: str
    activity_name: str
    share: Decimal


@dataclass(frozen=True)
class MappingTrace:
    relationships: tuple[MappingRelationship, ...]
    activities: tuple[tuple[str, str], ...]
    by_boq: tuple[tuple[str, tuple[int, ...]], ...]
    by_activity: tuple[tuple[str, tuple[int, ...]], ...]

    @property
    def capacity(self):
        return max((len(rows) for _, rows in self.by_boq + self.by_activity), default=0)


def mapping_trace(projection):
    """Indices reference one canonical edge list; payment amounts are never read."""
    names = dict(projection.activity_names)
    if len(names) != len(projection.activity_names):
        raise ValueError('Duplicate Activity identity in mapping trace')
    edges, forward, reverse = [], {}, {identity: [] for identity in names}
    seen = set()
    for entry in projection.items:
        item = entry.item
        if item.role != 'item':
            continue
        if item.key in forward:
            raise ValueError('Duplicate BOQ identity in mapping trace')
        forward[item.key] = []
        first = entry.periods[0].contributions if entry.periods else ()
        expected = tuple((c.activity_id, c.share) for c in first)
        if any(tuple((c.activity_id, c.share) for c in period.contributions) != expected
               for period in entry.periods):
            raise ValueError('Mapping relationships differ between periods')
        for contributor in first:
            identity, share = contributor.activity_id, contributor.share
            pair = item.key, identity
            if pair in seen or not isinstance(share, Decimal) or not share.is_finite() or not 0 < share <= 1:
                raise ValueError('Invalid or duplicate saved Mapping Share')
            seen.add(pair)
            # Older direct renderer callers may omit presentation names. Never
            # invent one; the service supplies validated current-main labels.
            name = names.get(identity, 'Name unavailable')
            index = len(edges)
            edges.append(MappingRelationship(item.key, item.description, identity, name, share))
            forward[item.key].append(index)
            reverse.setdefault(identity, []).append(index)
    activities = tuple((identity, names.get(identity, 'Name unavailable')) for identity in sorted(reverse))
    return MappingTrace(tuple(edges), activities,
                        tuple((key, tuple(rows)) for key, rows in forward.items()),
                        tuple((key, tuple(reverse[key])) for key, _ in activities))
