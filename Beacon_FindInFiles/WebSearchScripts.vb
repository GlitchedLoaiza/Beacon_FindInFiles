Imports System.Text.Json

Namespace Beacon
    Public NotInheritable Class WebSearchScripts
        Private Sub New()
        End Sub

        Public Shared Function Capture(maxCharacters As Integer, token As String) As String
            Return "(function() {
    document.querySelectorAll('mark.search-highlight,mark.current-highlight').forEach(function(mark) {
        var parent = mark.parentNode;
        parent.replaceChild(document.createTextNode(mark.textContent), mark);
        parent.normalize();
    });
    var nodes = [], parts = [], offset = 0;
    if (!document.body) return null;
    var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
        acceptNode: function(node) {
            var parent = node.parentElement;
            if (!parent || parent.closest('script,style,noscript') || !parent.getClientRects().length || getComputedStyle(parent).visibility === 'hidden') return NodeFilter.FILTER_REJECT;
            return NodeFilter.FILTER_ACCEPT;
        }
    });
    var node;
    while ((node = walker.nextNode()) && offset < __LIMIT__ && nodes.length < 50000) {
        var text = node.nodeValue.slice(0, __LIMIT__ - offset);
        nodes.push({node: node, original: node.nodeValue, start: offset, end: offset + text.length});
        parts.push(text);
        offset += text.length;
    }
    window.__beaconNodes = nodes;
    window.__beaconCapture = __TOKEN__;
    return parts.join('');
})();".Replace("__LIMIT__", maxCharacters.ToString(Globalization.CultureInfo.InvariantCulture)).Replace("__TOKEN__", JsonSerializer.Serialize(token))
        End Function

        Public Shared Function Apply(spans As IEnumerable(Of SearchSpan), token As String) As String
            Return "(function() {
    if (window.__beaconCapture !== __TOKEN__ || !window.__beaconNodes) return null;
    var ranges = __RANGES__, nodes = window.__beaconNodes, cursor = 0, ids = new Set();
    if (nodes.some(function(item) { return !item.node.isConnected || item.node.nodeValue !== item.original; })) return null;
    nodes.forEach(function(item) {
        var node = item.node;
        if (!node.parentNode) return;
        while (cursor < ranges.length && ranges[cursor].Start + ranges[cursor].Length <= item.start) cursor++;
        var fragment = document.createDocumentFragment(), position = 0, changed = false;
        for (var k = cursor; k < ranges.length && ranges[k].Start < item.end; k++) {
            var range = ranges[k], start = Math.max(range.Start, item.start) - item.start;
            var end = Math.min(range.Start + range.Length, item.end) - item.start;
            if (end <= start) continue;
            fragment.appendChild(document.createTextNode(node.nodeValue.slice(position, start)));
            var mark = document.createElement('mark');
            mark.className = 'search-highlight';
            mark.dataset.matchIndex = k;
            mark.textContent = node.nodeValue.slice(start, end);
            fragment.appendChild(mark);
            position = end;
            ids.add(k);
            changed = true;
        }
        if (changed) {
            fragment.appendChild(document.createTextNode(node.nodeValue.slice(position)));
            node.parentNode.replaceChild(fragment, node);
        }
    });
    var ordered = Array.from(ids).sort(function(a,b) { return a-b; }), compact = new Map();
    ordered.forEach(function(id,index) { compact.set(id,index); });
    document.querySelectorAll('mark.search-highlight').forEach(function(mark) { mark.dataset.matchIndex = compact.get(Number(mark.dataset.matchIndex)); });
    delete window.__beaconNodes;
    return ordered.length;
})();".Replace("__TOKEN__", JsonSerializer.Serialize(token)).Replace("__RANGES__", JsonSerializer.Serialize(spans))
        End Function
    End Class
End Namespace
