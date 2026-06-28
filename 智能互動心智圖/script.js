// Override d3.linkHorizontal to support toggling between straight lines and curves
(function () {
    const originalLinkHorizontal = d3.linkHorizontal;
    window.useStraightLines = true; // Default to straight lines with right angles

    d3.linkHorizontal = function () {
        const originalGenerator = originalLinkHorizontal();
        let sourceAccessor = d => d.source;
        let targetAccessor = d => d.target;
        let xAccessor = d => d[0];
        let yAccessor = d => d[1];

        const generator = function (d) {
            if (!window.useStraightLines) {
                return originalGenerator(d);
            }
            const s = sourceAccessor(d);
            const t = targetAccessor(d);
            const x0 = xAccessor(s);
            const y0 = yAccessor(s);
            const x1 = xAccessor(t);
            const y1 = yAccessor(t);
            const x_mid = (x0 + x1) / 2;
            return `M${x0},${y0} L${x_mid},${y0} L${x_mid},${y1} L${x1},${y1}`;
        };

        generator.source = function (_) {
            if (arguments.length) {
                sourceAccessor = _;
                originalGenerator.source(_);
                return generator;
            }
            return sourceAccessor;
        };
        generator.target = function (_) {
            if (arguments.length) {
                targetAccessor = _;
                originalGenerator.target(_);
                return generator;
            }
            return targetAccessor;
        };
        generator.x = function (_) {
            if (arguments.length) {
                xAccessor = _;
                originalGenerator.x(_);
                return generator;
            }
            return xAccessor;
        };
        generator.y = function (_) {
            if (arguments.length) {
                yAccessor = _;
                originalGenerator.y(_);
                return generator;
            }
            return yAccessor;
        };

        return generator;
    };
})();

var transformer, mm;
let activeNodeData = null;
let activeNodeEl = null;

// Simple markdown to tree parser (fallback if Transformer unavailable)
function parseMarkdown(md) {
    const lines = md.split('\n');
    const root = { content: '', children: [], depth: 0 };
    const stack = [root];

    for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed) continue;

        let depth, content;
        const headingMatch = trimmed.match(/^(#{1,6})\s+(.*)/);
        const bulletMatch = trimmed.match(/^[-*]\s+(.*)/);
        const indentedBulletMatch = line.match(/^(\s+)[-*]\s+(.*)/);

        if (headingMatch) {
            depth = headingMatch[1].length;
            content = headingMatch[2];
        } else if (indentedBulletMatch) {
            const indent = indentedBulletMatch[1].length;
            depth = 7 + Math.floor(indent / 2);
            content = indentedBulletMatch[2];
        } else if (bulletMatch) {
            depth = 7;
            content = bulletMatch[1];
        } else {
            continue;
        }

        const node = { content, children: [], depth };

        while (stack.length > 1 && stack[stack.length - 1].depth >= depth) {
            stack.pop();
        }
        stack[stack.length - 1].children.push(node);
        stack.push(node);
    }

    return root.children.length === 1 ? root.children[0] : root;
}

function buildMarkmapData(node) {
    return {
        content: node.content,
        children: (node.children || []).map(buildMarkmapData)
    };
}

function alignNodesWithLines(root, markdown) {
    const lines = markdown.split('\n');
    const nodeLines = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();
        if (!trimmed) continue;

        const isHeading = trimmed.match(/^(#{1,6})\s+/);
        const isBullet = trimmed.match(/^[-*]\s+/);
        const isIndentedBullet = line.match(/^(\s+)[-*]\s+/);

        if (isHeading || isBullet || isIndentedBullet) {
            nodeLines.push({
                lineIndex: i,
                lineContent: line
            });
        }
    }

    const treeNodes = [];
    function traverse(node) {
        if (!node) return;
        if (node.depth === 0 && !node.content && node.children && node.children.length > 0) {
            // Virtual root
        } else {
            treeNodes.push(node);
        }
        if (node.children) {
            node.children.forEach(traverse);
        }
    }

    traverse(root);

    const limit = Math.min(treeNodes.length, nodeLines.length);
    for (let i = 0; i < limit; i++) {
        treeNodes[i].lineIndex = nodeLines[i].lineIndex;
    }
}

const darkColors = ['#818cf8', '#c084fc', '#2dd4bf', '#fbbf24', '#f472b6', '#60a5fa', '#f87171', '#a3e635', '#22d3ee'];
const lightColors = ['#4f46e5', '#9333ea', '#0d9488', '#d97706', '#db2777', '#2563eb', '#dc2626', '#65a30d', '#0891b2'];

const darkStyle = (id) => `
    .${id} g { fill: #ffffff; font-weight: 500; font-family: ui-sans-serif, system-ui, sans-serif; }
    .${id} foreignObject { color: #ffffff; }
    .${id} circle { fill: #0f172a; }
`;

const lightStyle = (id) => `
    .${id} g { fill: #0f172a; font-weight: 500; font-family: ui-sans-serif, system-ui, sans-serif; }
    .${id} foreignObject { color: #0f172a; }
    .${id} circle { fill: #ffffff; }
`;

let isAllExpanded = false;

function setAllNodesFold(node, foldVal) {
    if (!node) return;
    if (!node.payload) node.payload = {};
    node.payload.fold = foldVal;
    if (node.children) {
        node.children.forEach(child => setAllNodesFold(child, foldVal));
    }
}

function collapseAllNodes(node, depth = 1) {
    if (!node) return;
    if (!node.payload) node.payload = {};
    if (depth > 2) {
        node.payload.fold = 1;
    } else {
        delete node.payload.fold;
    }
    if (node.children) {
        node.children.forEach(child => collapseAllNodes(child, depth + 1));
    }
}

function resetExpandButton() {
    const btn = document.querySelector('#btn-expand-all');
    const mobileBtn = document.querySelector('#mobile-btn-expand-all');

    const updateBtnState = (b) => {
        if (!b) return;
        const icon = b.querySelector('i');
        const span = b.querySelector('span');
        icon.className = "fa-solid fa-angles-down";
        span.innerText = "一鍵展開";
    };

    updateBtnState(btn);
    updateBtnState(mobileBtn);
    isAllExpanded = false;
}

function toggleExpandAll() {
    const btn = document.querySelector('#btn-expand-all');
    const mobileBtn = document.querySelector('#mobile-btn-expand-all');

    const setExpandState = (b, expanded) => {
        if (!b) return;
        const icon = b.querySelector('i');
        const span = b.querySelector('span');
        if (expanded) {
            icon.className = "fa-solid fa-rotate-left";
            span.innerText = "恢復預設";
        } else {
            icon.className = "fa-solid fa-angles-down";
            span.innerText = "一鍵展開";
        }
    };

    if (isAllExpanded) {
        collapseAllNodes(mm.state.data);
        mm.setData(mm.state.data);
        setExpandState(btn, false);
        setExpandState(mobileBtn, false);
        isAllExpanded = false;
    } else {
        setAllNodesFold(mm.state.data, 0);
        mm.renderData();
        setExpandState(btn, true);
        setExpandState(mobileBtn, true);
        isAllExpanded = true;
    }
}

function toggleLineStyle() {
    const btn = document.querySelector('#btn-toggle-line');
    const mobileBtn = document.querySelector('#mobile-btn-toggle-line');

    window.useStraightLines = !window.useStraightLines;

    const updateStyleState = (b) => {
        if (!b) return;
        const icon = b.querySelector('i');
        const span = b.querySelector('span');
        if (window.useStraightLines) {
            icon.className = "fa-solid fa-route";
            span.innerText = "直角折線";
        } else {
            icon.className = "fa-solid fa-bezier-curve";
            span.innerText = "曲線連結";
        }
    };

    updateStyleState(btn);
    updateStyleState(mobileBtn);

    if (mm) {
        mm.renderData();
    }
}

function decodeHTMLEntities(text) {
    const textArea = document.createElement('textarea');
    textArea.innerHTML = text;
    return textArea.value;
}

let lastActiveNode = null;

function findActiveNodeForLine(lineIndex) {
    if (!mm || !mm.state.data) return null;
    let closestNode = null;
    let maxLine = -1;

    function walk(node) {
        if (!node) return;
        if (node.lineIndex !== undefined && node.lineIndex <= lineIndex && node.lineIndex > maxLine) {
            maxLine = node.lineIndex;
            closestNode = node;
        }
        if (node.children) {
            node.children.forEach(walk);
        }
    }

    walk(mm.state.data);
    return closestNode;
}

function unfoldAncestors(targetNode) {
    if (!mm || !mm.state.data) return false;
    let path = [];
    let found = false;

    function findPath(node, currentPath) {
        if (!node || found) return;
        currentPath.push(node);
        if (node === targetNode) {
            path = [...currentPath];
            found = true;
            return;
        }
        if (node.children) {
            node.children.forEach(child => findPath(child, currentPath));
        }
        currentPath.pop();
    }

    findPath(mm.state.data, []);

    let updated = false;
    for (let i = 0; i < path.length - 1; i++) {
        if (!path[i].payload) path[i].payload = {};
        if (path[i].payload.fold !== 0) {
            path[i].payload.fold = 0;
            updated = true;
        }
    }
    return updated;
}

function syncMapWithEditorCursor() {
    if (!mm || !mm.state.data) return;
    const editor = document.querySelector('#editor');
    const cursorPos = editor.selectionStart;
    const textBefore = editor.value.substring(0, cursorPos);
    const lineIndex = textBefore.split('\n').length - 1;

    const activeNode = findActiveNodeForLine(lineIndex);
    if (!activeNode) {
        if (lastActiveNode !== null) {
            lastActiveNode = null;
            mm.setHighlight(null);
        }
        return;
    }

    if (activeNode !== lastActiveNode) {
        lastActiveNode = activeNode;
        const wasFolded = unfoldAncestors(activeNode);
        if (wasFolded) {
            mm.renderData().then(() => {
                mm.setHighlight(activeNode);
                mm.centerNode(activeNode);
            });
        } else {
            mm.setHighlight(activeNode);
            mm.centerNode(activeNode);
        }
    }
}

function initMarkmap() {
    const { Markmap } = window.markmap;
    const svgEl = document.querySelector('#markmap');
    const isDark = document.documentElement.classList.contains('dark');
    const currentColors = isDark ? darkColors : lightColors;
    const currentStyle = isDark ? darkStyle : lightStyle;

    mm = Markmap.create(svgEl, {
        zoom: true,
        pan: true,
        initialExpandLevel: 2,
        spacingVertical: 15,
        color: (node) => {
            const depth = node.state ? node.state.depth : (node.depth || 0);
            return currentColors[depth % currentColors.length];
        },
        style: currentStyle
    });

    // Try to get Transformer from markmap-lib
    if (window.markmap && window.markmap.Transformer) {
        transformer = new window.markmap.Transformer();
        console.log('Using markmap Transformer');
    } else {
        console.log('Transformer not available, using built-in parser');
    }

    setTimeout(() => { updateMarkmap(true); }, 100);
    setupEditEvents();
}

function toggleTheme() {
    const html = document.documentElement;
    const isDark = html.classList.toggle('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
    if (mm) {
        const currentColors = isDark ? darkColors : lightColors;
        const currentStyle = isDark ? darkStyle : lightStyle;
        mm.setOptions({
            color: (node) => {
                const depth = node.state ? node.state.depth : (node.depth || 0);
                return currentColors[depth % currentColors.length];
            },
            style: currentStyle
        });
        mm.updateStyle();
        mm.renderData();
    }
}

function saveFoldState(node, map = {}) {
    if (!node) return map;
    if (node.content !== undefined) {
        const decodedKey = decodeHTMLEntities(node.content);
        map[decodedKey] = node.payload?.fold ?? null;
    }
    (node.children || []).forEach(c => saveFoldState(c, map));
    return map;
}

function restoreFoldState(node, map) {
    if (!node) return;
    if (node.content !== undefined) {
        const decodedKey = decodeHTMLEntities(node.content);
        if (map[decodedKey] !== undefined && map[decodedKey] !== null) {
            if (!node.payload) node.payload = {};
            node.payload.fold = map[decodedKey];
        }
    }
    (node.children || []).forEach(c => restoreFoldState(c, map));
}

function updateMarkmap(isFirstLoad = false) {
    if (!mm) return;
    const markdown = document.querySelector('#editor').value;

    // 編輯時快照現有折疊狀態，稍後還原
    const foldSnapshot = (!isFirstLoad && mm.state.data) ? saveFoldState(mm.state.data) : {};

    let root;
    if (transformer) {
        const result = transformer.transform(markdown);
        root = result.root;
    } else {
        root = buildMarkmapData(parseMarkdown(markdown));
    }

    alignNodesWithLines(root, markdown);
    lastActiveNode = null;

    // 將快照的折疊狀態還原到新 root（首次載入時略過，使用預設 initialExpandLevel）
    if (!isFirstLoad) {
        restoreFoldState(root, foldSnapshot);
    }

    mm.setData(root);

    // 只在首次載入時 fit，避免編輯時視角跳動
    if (isFirstLoad) mm.fit();

    resetExpandButton();
}

function setupEditEvents() {
    const svgEl = document.querySelector('#markmap');
    const contextMenu = document.querySelector('#context-menu');
    const inlineEditorContainer = document.querySelector('#inline-editor-container');
    const inlineEditor = document.querySelector('#inline-editor');

    svgEl.addEventListener('contextmenu', (e) => {
        const nodeEl = e.target.closest('.markmap-node');
        if (!nodeEl) return;
        e.preventDefault();
        e.stopPropagation();

        activeNodeEl = nodeEl;
        activeNodeData = nodeEl.__data__;

        if (!activeNodeData || activeNodeData.lineIndex === undefined) return;

        const containerRect = document.querySelector('#map-container').getBoundingClientRect();
        contextMenu.style.left = `${e.clientX - containerRect.left}px`;
        contextMenu.style.top = `${e.clientY - containerRect.top}px`;
        contextMenu.classList.remove('hidden');
        inlineEditorContainer.classList.add('hidden');
    });

    svgEl.addEventListener('dblclick', (e) => {
        const nodeEl = e.target.closest('.markmap-node');
        if (!nodeEl) return;
        e.preventDefault();
        e.stopPropagation();

        activeNodeEl = nodeEl;
        activeNodeData = nodeEl.__data__;

        if (!activeNodeData || activeNodeData.lineIndex === undefined) return;

        showInlineEditor();
    }, true);

    document.addEventListener('click', (e) => {
        if (!contextMenu.contains(e.target)) {
            contextMenu.classList.add('hidden');
        }
        if (!inlineEditorContainer.contains(e.target) && !e.target.closest('.markmap-node')) {
            commitInlineEdit();
        }
    });

    inlineEditor.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            commitInlineEdit();
        } else if (e.key === 'Escape') {
            inlineEditorContainer.classList.add('hidden');
        }
    });

    const editor = document.querySelector('#editor');
    editor.addEventListener('click', syncMapWithEditorCursor);
    editor.addEventListener('keyup', (e) => {
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown'].includes(e.key)) {
            syncMapWithEditorCursor();
        }
    });
    editor.addEventListener('input', () => {
        clearTimeout(timeout);
        timeout = setTimeout(() => updateMarkmap(false), 300);
    });

    svgEl.addEventListener('click', (event) => {
        const nodeEl = event.target.closest('.markmap-node');
        if (nodeEl) {
            const d = nodeEl.__data__;
            if (d && d.lineIndex !== undefined) {
                const lines = editor.value.split('\n');
                let charIndex = 0;
                for (let i = 0; i < d.lineIndex; i++) {
                    charIndex += lines[i].length + 1;
                }
                editor.focus();
                editor.setSelectionRange(charIndex, charIndex + lines[d.lineIndex].length);

                const lineHeight = 24;
                editor.scrollTop = d.lineIndex * lineHeight - editor.clientHeight / 2;

                mm.setHighlight(d);
            }
        }
    }, true);
}

function showInlineEditor() {
    const inlineEditorContainer = document.querySelector('#inline-editor-container');
    const inlineEditor = document.querySelector('#inline-editor');
    const containerRect = document.querySelector('#map-container').getBoundingClientRect();
    const nodeRect = activeNodeEl.getBoundingClientRect();

    inlineEditor.value = decodeHTMLEntities(activeNodeData.content);
    inlineEditorContainer.style.left = `${nodeRect.left - containerRect.left}px`;
    inlineEditorContainer.style.top = `${nodeRect.top - containerRect.top - 5}px`;
    inlineEditorContainer.style.width = `${Math.max(nodeRect.width + 10, 160)}px`;
    inlineEditorContainer.classList.remove('hidden');
    inlineEditor.focus();
    inlineEditor.select();

    document.querySelector('#context-menu').classList.add('hidden');
}

function commitInlineEdit() {
    const inlineEditorContainer = document.querySelector('#inline-editor-container');
    if (inlineEditorContainer.classList.contains('hidden') || !activeNodeData) return;

    const newValue = document.querySelector('#inline-editor').value.trim();
    if (newValue && newValue !== decodeHTMLEntities(activeNodeData.content)) {
        updateMarkdownLine(activeNodeData.lineIndex, newValue);
    }
    inlineEditorContainer.classList.add('hidden');
}

function getLineDepth(line) {
    const trimmed = line.trim();
    const headingMatch = trimmed.match(/^(#{1,6})\s+/);
    const indentedBulletMatch = line.match(/^(\s+)[-*]\s+/);
    const bulletMatch = trimmed.match(/^[-*]\s+/);

    if (headingMatch) {
        return headingMatch[1].length;
    } else if (indentedBulletMatch) {
        const indent = indentedBulletMatch[1].length;
        return 7 + Math.floor(indent / 2);
    } else if (bulletMatch) {
        return 7;
    }
    return 999;
}

function findDescendantEndIndex(lines, startIndex, baseDepth) {
    let i = startIndex + 1;
    while (i < lines.length) {
        const line = lines[i];
        if (!line.trim()) {
            i++;
            continue;
        }
        const depth = getLineDepth(line);
        if (depth <= baseDepth) {
            break;
        }
        i++;
    }
    return i;
}

function updateMarkdownLine(lineIndex, newText) {
    const editor = document.querySelector('#editor');
    const lines = editor.value.split('\n');
    const originalLine = lines[lineIndex];
    const match = originalLine.match(/^(\s*(?:#{1,6}|[-*])\s+)/);
    const prefix = match ? match[1] : '';

    lines[lineIndex] = prefix + newText;
    editor.value = lines.join('\n');
    updateMarkmap(false);
}

function menuEditNode() {
    document.querySelector('#context-menu').classList.add('hidden');
    showInlineEditor();
}

function menuAddChild() {
    document.querySelector('#context-menu').classList.add('hidden');
    if (!activeNodeData) return;

    const editor = document.querySelector('#editor');
    const lines = editor.value.split('\n');
    const lineIndex = activeNodeData.lineIndex;
    const parentLine = lines[lineIndex];

    const headingMatch = parentLine.match(/^(\s*)(#{1,6})\s+/);
    const bulletMatch = parentLine.match(/^(\s*)([-*])\s+/);
    let childPrefix = '';

    if (headingMatch) {
        const indent = headingMatch[1];
        const hashes = headingMatch[2];
        if (hashes.length < 6) {
            childPrefix = indent + hashes + '# ';
        } else {
            childPrefix = indent + '- ';
        }
    } else if (bulletMatch) {
        const indent = bulletMatch[1];
        childPrefix = indent + '  - ';
    } else {
        childPrefix = '- ';
    }

    lines.splice(lineIndex + 1, 0, childPrefix + '新子節點');
    editor.value = lines.join('\n');
    updateMarkmap(false);

    setTimeout(() => {
        const svgEl = document.querySelector('#markmap');
        const nodeEls = svgEl.querySelectorAll('.markmap-node');
        for (const el of nodeEls) {
            if (el.__data__ && el.__data__.lineIndex === lineIndex + 1) {
                activeNodeEl = el;
                activeNodeData = el.__data__;
                showInlineEditor();
                break;
            }
        }
    }, 400);
}

function menuAddSibling() {
    document.querySelector('#context-menu').classList.add('hidden');
    if (!activeNodeData) return;

    const editor = document.querySelector('#editor');
    const lines = editor.value.split('\n');
    const lineIndex = activeNodeData.lineIndex;
    const currentLine = lines[lineIndex];
    const baseDepth = getLineDepth(currentLine);

    const match = currentLine.match(/^(\s*(?:#{1,6}|[-*])\s+)/);
    const prefix = match ? match[1] : '- ';

    const insertIndex = findDescendantEndIndex(lines, lineIndex, baseDepth);

    lines.splice(insertIndex, 0, prefix + '新節點');
    editor.value = lines.join('\n');
    updateMarkmap(false);

    setTimeout(() => {
        const svgEl = document.querySelector('#markmap');
        const nodeEls = svgEl.querySelectorAll('.markmap-node');
        for (const el of nodeEls) {
            if (el.__data__ && el.__data__.lineIndex === insertIndex) {
                activeNodeEl = el;
                activeNodeData = el.__data__;
                showInlineEditor();
                break;
            }
        }
    }, 400);
}

function menuDeleteNode() {
    document.querySelector('#context-menu').classList.add('hidden');
    if (!activeNodeData) return;

    const editor = document.querySelector('#editor');
    const lines = editor.value.split('\n');
    const lineIndex = activeNodeData.lineIndex;
    const currentLine = lines[lineIndex];
    const baseDepth = getLineDepth(currentLine);

    const endIndex = findDescendantEndIndex(lines, lineIndex, baseDepth);
    lines.splice(lineIndex, endIndex - lineIndex);

    editor.value = lines.join('\n');
    updateMarkmap(false);
}

let timeout;

let lastSplitPercent = 50;
let lastSplitPercentMobile = 50;
let activeViewMode = 'split';

function isMobileView() {
    return window.innerWidth < 768;
}

function toggleMobileMenu() {
    const panel = document.querySelector('#mobile-menu-panel');
    const backdrop = document.querySelector('#mobile-menu-backdrop');
    if (!panel || !backdrop) return;

    if (panel.classList.contains('hidden')) {
        panel.classList.remove('hidden');
        backdrop.classList.remove('hidden');
        setTimeout(() => {
            panel.classList.remove('translate-x-full');
            backdrop.classList.remove('opacity-0');
            backdrop.classList.add('opacity-100');
        }, 10);
    } else {
        panel.classList.add('translate-x-full');
        backdrop.classList.remove('opacity-100');
        backdrop.classList.add('opacity-0');
        setTimeout(() => {
            panel.classList.add('hidden');
            backdrop.classList.add('hidden');
        }, 300);
    }
}

function setupDragBar() {
    const dragBar = document.querySelector('#drag-bar');
    const editorCol = document.querySelector('#editor-container');
    const mapCol = document.querySelector('#map-container');
    const main = document.querySelector('main');

    if (!dragBar) return;

    let isDragging = false;

    function onDragStart(e) {
        isDragging = true;
        mapCol.style.pointerEvents = 'none';
        document.body.style.cursor = isMobileView() ? 'row-resize' : 'col-resize';
        document.body.style.userSelect = 'none';
    }

    function onDragMove(e) {
        if (!isDragging) return;
        if (e.cancelable) e.preventDefault();

        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;

        const mainRect = main.getBoundingClientRect();

        if (isMobileView()) {
            const offsetY = clientY - mainRect.top;
            const percentage = (offsetY / mainRect.height) * 100;
            if (percentage >= 15 && percentage <= 85) {
                lastSplitPercentMobile = percentage;
                editorCol.style.height = `${percentage}%`;
                editorCol.style.width = '100%';
                mapCol.style.height = `${100 - percentage}%`;
                mapCol.style.width = '100%';
            }
        } else {
            const offsetX = clientX - mainRect.left;
            const percentage = (offsetX / mainRect.width) * 100;
            if (percentage >= 20 && percentage <= 80) {
                lastSplitPercent = percentage;
                editorCol.style.width = `${percentage}%`;
                editorCol.style.height = '100%';
                mapCol.style.width = `${100 - percentage}%`;
                mapCol.style.height = '100%';
            }
        }
    }

    function onDragEnd() {
        if (isDragging) {
            isDragging = false;
            mapCol.style.pointerEvents = 'auto';
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            if (mm) mm.fit();
        }
    }

    // Mouse Events
    dragBar.addEventListener('mousedown', onDragStart);
    document.addEventListener('mousemove', onDragMove);
    document.addEventListener('mouseup', onDragEnd);

    // Touch Events (Mobile)
    dragBar.addEventListener('touchstart', onDragStart, { passive: true });
    document.addEventListener('touchmove', onDragMove, { passive: false });
    document.addEventListener('touchend', onDragEnd);
}

function switchView(mode) {
    activeViewMode = mode;
    const editorCol = document.querySelector('#editor-container');
    const mapCol = document.querySelector('#map-container');
    const dragBar = document.querySelector('#drag-bar');

    // Sync buttons for desktop
    const btns = {
        split: document.querySelector('#btn-view-split'),
        edit: document.querySelector('#btn-view-edit'),
        map: document.querySelector('#btn-view-map')
    };
    Object.values(btns).forEach(btn => {
        if (btn) btn.className = "px-3 py-1.5 rounded-md text-xs font-medium transition-all text-slate-500 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white flex items-center space-x-1";
    });

    // Sync buttons for mobile
    const mobileBtns = {
        split: document.querySelector('#mobile-btn-view-split'),
        edit: document.querySelector('#mobile-btn-view-edit'),
        map: document.querySelector('#mobile-btn-view-map')
    };
    Object.values(mobileBtns).forEach(btn => {
        if (btn) btn.className = "py-1.5 rounded-md text-xs font-medium text-slate-500 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white flex flex-col items-center justify-center";
    });

    if (mode === 'split') {
        if (btns.split) btns.split.className = "px-3 py-1.5 rounded-md text-xs font-medium transition-all bg-indigo-600 text-white shadow-sm flex items-center space-x-1";
        if (mobileBtns.split) mobileBtns.split.className = "py-1.5 rounded-md text-xs font-medium bg-indigo-600 text-white shadow-sm flex flex-col items-center justify-center";

        editorCol.style.display = "flex";
        mapCol.style.display = "flex";
        if (dragBar) dragBar.style.display = "flex";

        if (isMobileView()) {
            editorCol.style.width = '100%';
            editorCol.style.height = `${lastSplitPercentMobile}%`;
            mapCol.style.width = '100%';
            mapCol.style.height = `${100 - lastSplitPercentMobile}%`;
        } else {
            editorCol.style.width = `${lastSplitPercent}%`;
            editorCol.style.height = '100%';
            mapCol.style.width = `${100 - lastSplitPercent}%`;
            mapCol.style.height = '100%';
        }
    } else if (mode === 'edit') {
        if (btns.edit) btns.edit.className = "px-3 py-1.5 rounded-md text-xs font-medium transition-all bg-indigo-600 text-white shadow-sm flex items-center space-x-1";
        if (mobileBtns.edit) mobileBtns.edit.className = "py-1.5 rounded-md text-xs font-medium bg-indigo-600 text-white shadow-sm flex flex-col items-center justify-center";

        editorCol.style.display = "flex";
        editorCol.style.width = "100%";
        editorCol.style.height = "100%";
        mapCol.style.display = "none";
        if (dragBar) dragBar.style.display = "none";
    } else if (mode === 'map') {
        if (btns.map) btns.map.className = "px-3 py-1.5 rounded-md text-xs font-medium transition-all bg-indigo-600 text-white shadow-sm flex items-center space-x-1";
        if (mobileBtns.map) mobileBtns.map.className = "py-1.5 rounded-md text-xs font-medium bg-indigo-600 text-white shadow-sm flex flex-col items-center justify-center";

        editorCol.style.display = "none";
        mapCol.style.display = "flex";
        mapCol.style.width = "100%";
        mapCol.style.height = "100%";
        if (dragBar) dragBar.style.display = "none";
    }
    setTimeout(() => { if (mm && mode !== 'edit') { mm.fit(); } }, 300);
}

window.addEventListener('resize', () => {
    if (activeViewMode === 'split') {
        const editorCol = document.querySelector('#editor-container');
        const mapCol = document.querySelector('#map-container');
        if (isMobileView()) {
            editorCol.style.width = '100%';
            editorCol.style.height = `${lastSplitPercentMobile}%`;
            mapCol.style.width = '100%';
            mapCol.style.height = `${100 - lastSplitPercentMobile}%`;
        } else {
            editorCol.style.width = `${lastSplitPercent}%`;
            editorCol.style.height = '100%';
            mapCol.style.width = `${100 - lastSplitPercent}%`;
            mapCol.style.height = '100%';
        }
    }
    if (mm && document.querySelector('#map-container').style.display !== 'none') { mm.fit(); }
});

function zoomIn() { mm.rescale(1.25); }
function zoomOut() { mm.rescale(0.8); }
function zoomFit() { mm.fit(); }

function exportMarkdown() {
    const text = document.querySelector('#editor').value;
    const blob = new Blob([text], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '陳仲豪行銷課程大綱.md';
    a.click();
    URL.revokeObjectURL(url);
}

function exportSVG() {
    const svgEl = document.querySelector('#markmap');
    const clone = svgEl.cloneNode(true);
    const g = svgEl.querySelector('g');
    const bbox = g.getBBox();
    clone.setAttribute('viewBox', `${bbox.x - 50} ${bbox.y - 50} ${bbox.width + 100} ${bbox.height + 100}`);
    clone.setAttribute('width', bbox.width + 100);
    clone.setAttribute('height', bbox.height + 100);
    clone.style.backgroundColor = document.documentElement.classList.contains('dark') ? '#0f172a' : '#f8fafc';
    const serializer = new XMLSerializer();
    let source = serializer.serializeToString(clone);
    if (!source.match(/^<\?xml/)) {
        source = '<?xml version="1.0" standalone="no"?>\r\n' + source;
    }
    const blob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '陳仲豪行銷課程大綱_心智圖.svg';
    a.click();
    URL.revokeObjectURL(url);
}

function formatEditorHeadings() {
    const editor = document.querySelector('#editor');
    if (!editor) return;
    const lines = editor.value.split('\n');
    const formatted = [];
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.trim().startsWith('#') && i > 0 && lines[i - 1].trim() !== '') {
            formatted.push('');
        }
        formatted.push(line);
    }
    editor.value = formatted.join('\n');
}

function toggleInfoTooltip() {
    const tooltip = document.querySelector('#info-tooltip');
    if (tooltip) {
        tooltip.classList.toggle('hidden');
    }
}

// 點擊空白處自動收合提示
document.addEventListener('click', (e) => {
    const container = document.querySelector('#btn-info-toggle')?.parentElement;
    const tooltip = document.querySelector('#info-tooltip');
    if (container && tooltip && !container.contains(e.target) && !tooltip.classList.contains('hidden')) {
        tooltip.classList.add('hidden');
    }
});

window.onload = () => {
    formatEditorHeadings();
    initMarkmap();
    setupDragBar();
};
