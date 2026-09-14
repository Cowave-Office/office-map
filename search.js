window.OfficeMapSearch = (() => {
  const compact = value => String(value || '').normalize('NFKC').replace(/\s+/g, '');
  const normalize = value => compact(value).toLowerCase();
  const display = value => String(value || '').trim().replace(/\s+/g, ' ');
  const facilities = /회의|화장|휴게|서버|탕비|교육장|출입|창고|사물함|비품|대표실|임원실|세미나|샤워|복합기|엘리베이터|카페테리아|라운지|공석|외부인력|망분리/i;
  const organization = value => {
    const text = compact(value);
    return !facilities.test(text) && (/(팀|본부|파트|그룹|담당|사업부|센터|셀)$/.test(text)
      || (text.length >= 4 && /실$/.test(text)) || ['QA', '경영지원', '서비스UX/UI'].includes(text));
  };
  const person = value => /^[가-힣]{2,4}[a-zA-Z]{0,3}$/.test(compact(value)) && !facilities.test(value) && !organization(value);
  const gap = (a, lengthA, b, lengthB) => Math.max(0, a - b - lengthB, b - a - lengthA);
  const seatColor = item => item.bg && !/^(#?(ffffff|000000)|theme:0:0\.0|rgb:FF(FFFFFF|000000):0\.0)$/i.test(item.bg) ? item.bg.toLowerCase() : '';

  function teamFor(item, headings) {
    // The workbook is a floor plan, not an HR roster. Ambiguous areas stay unassigned.
    const ranked = headings.map(({ header, colors }) => {
      const dx = gap(item.c, item.cs, header.c, header.cs);
      const dy = gap(item.r, item.rs, header.r, header.rs);
      const broad = /(본부|담당|사업부)$/.test(compact(header.v));
      const color = seatColor(item);
      const matchesColor = Boolean(color && colors.has(color));
      const conflictsColor = Boolean(color && colors.size && !matchesColor);
      return { header, dx, dy, matchesColor, conflictsColor, score: dx * (matchesColor ? 1.5 : 4) + dy + (broad ? 4 : 0) };
    }).filter(row => !row.conflictsColor && row.dx <= (row.matchesColor ? 12 : 3) && row.dy <= 18 && row.score <= 22)
      .sort((a, b) => a.score - b.score);
    if (!ranked.length) return '';
    const best = ranked[0];
    const competing = ranked.find(row => normalize(row.header.v) !== normalize(best.header.v));
    if (competing && competing.score - best.score < 2) return '';
    return display(best.header.v);
  }

  function buildIndex(items) {
    const covered = new Set();
    for (const it of items) {
      if (it.rs <= 1 && it.cs <= 1) continue;
      for (let r = it.r; r < it.r + it.rs; r++) {
        for (let c = it.c; c < it.c + it.cs; c++) {
          if (r !== it.r || c !== it.c) covered.add(`${r},${c}`);
        }
      }
    }
    const visible = items.filter(it => it.v && !covered.has(`${it.r},${it.c}`));
    const seats = visible.filter(it => person(it.v));
    const headings = visible.filter(it => organization(it.v)).map(header => {
      const adjacent = seats.map(item => ({ item, distance: gap(item.c, item.cs, header.c, header.cs) * 4 + gap(item.r, item.rs, header.r, header.rs) }))
        .filter(row => row.distance <= 3).sort((a, b) => a.distance - b.distance);
      const colors = new Set(adjacent.filter(row => row.distance === adjacent[0]?.distance).map(row => seatColor(row.item)).filter(Boolean));
      return { header, colors };
    });
    return visible.map(item => {
      const isPerson = person(item.v);
      return { item, text: normalize(item.v), name: isPerson ? compact(item.v) : display(item.v),
        isPerson, teamName: isPerson ? teamFor(item, headings) : '', isOrganization: organization(item.v) };
    });
  }

  function findMatches(indexes, value, currentTab, scope = null) {
    const query = normalize(value);
    if (!query) return [];
    const matches = [];
    for (const [tab, entries] of indexes) {
      if (scope && tab !== scope) continue;
      for (const entry of entries) {
        if (scope ? entry.text === query : entry.text.includes(query)) matches.push({ ...entry, tab });
      }
    }
    const rank = match => match.text === query ? 0
      : match.isPerson && match.text.replace(/[a-z]+$/, '') === query ? 1
      : match.text.startsWith(query) ? 2 : 3;
    return matches.sort((a, b) => rank(a) - rank(b) || a.name.localeCompare(b.name, 'ko', { numeric: true })
      || Number(b.tab === currentTab) - Number(a.tab === currentTab)
      || a.tab.localeCompare(b.tab) || a.item.r - b.item.r || a.item.c - b.item.c);
  }

  return { normalize, buildIndex, findMatches };
})();
