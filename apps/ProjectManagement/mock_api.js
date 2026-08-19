window.pywebview = {
    api: new Proxy({}, {
        get(target, prop) {
            if (prop === 'get_checklists') return async () => ({ checklists: [] });
            if (prop === 'get_timesheets') return async () => ({ weeks: {} });
            if (prop === 'get_templates') return async () => ({ templates: [] });
            if (prop === 'get_settings') return async () => ({});
            if (prop === 'get_version_info') return async () => ({ current_version: '1.0' });
            if (prop === 'get_user_settings') return async () => ({});
            if (prop === 'get_cad_commands') return async () => ({ commands: [] });
            if (prop === 'get_projects') return async () => [];
            if (prop === 'get_plugins') return async () => [];
            if (prop === 'init_app') return async () => ({ projects: [] });
            if (prop === 'select_files') return async (options = {}) => {
                const fileTypes = Array.isArray(options.file_types) ? options.file_types.join(' ') : '';
                if (/PDF Files/i.test(fileTypes)) {
                    return { status: 'success', paths: ['Symbol Counter Demo.pdf'] };
                }
                return { status: 'cancelled', paths: [] };
            };
            if (prop === 'prepare_symbol_count_documents') return async () => ({
                status: 'success',
                data: {
                    sessionId: 'preview-symbol-session',
                    pageCount: 2,
                    renderDpi: 120,
                    documents: [{
                        index: 0,
                        name: 'Symbol Counter Demo.pdf',
                        path: 'Symbol Counter Demo.pdf',
                        pageCount: 2,
                        pageIds: ['d1-p1', 'd1-p2']
                    }],
                    pages: [
                        { id: 'd1-p1', documentIndex: 0, pageIndex: 0, documentName: 'Symbol Counter Demo.pdf', label: 'E1.0', pixelWidth: 1600, pixelHeight: 1067 },
                        { id: 'd1-p2', documentIndex: 0, pageIndex: 1, documentName: 'Symbol Counter Demo.pdf', label: 'E2.0', pixelWidth: 1600, pixelHeight: 1067 }
                    ]
                }
            });
            if (prop === 'get_symbol_count_page') return async (_sessionId, pageId) => ({
                status: 'success',
                data: {
                    page: { id: pageId },
                    imageDataUrl: pageId === 'd1-p2' ? 'tmp/pdfs/arch_p4.png' : 'tmp/pdfs/electrical_p1.png'
                }
            });
            if (prop === 'count_pdf_symbols') return async (_sessionId, request = {}) => ({
                status: 'success',
                data: {
                    sourcePageId: request.sourcePageId || 'd1-p1',
                    templateRect: { x: 0.18, y: 0.21, width: 0.028, height: 0.042, pixelWidth: 45, pixelHeight: 45 },
                    threshold: request.threshold || 0.82,
                    total: 7,
                    pageCounts: [{ pageId: 'd1-p1', count: 5 }, { pageId: 'd1-p2', count: 2 }],
                    matches: [
                        { id: 'match-1', pageId: 'd1-p1', x: 0.18, y: 0.21, width: 0.028, height: 0.042, score: 0.99, rotation: 0, scale: 1, manual: false },
                        { id: 'match-2', pageId: 'd1-p1', x: 0.34, y: 0.34, width: 0.028, height: 0.042, score: 0.94, rotation: 90, scale: 1, manual: false },
                        { id: 'match-3', pageId: 'd1-p1', x: 0.49, y: 0.46, width: 0.028, height: 0.042, score: 0.91, rotation: 0, scale: 1, manual: false },
                        { id: 'match-4', pageId: 'd1-p1', x: 0.65, y: 0.58, width: 0.028, height: 0.042, score: 0.88, rotation: 270, scale: 1, manual: false },
                        { id: 'match-5', pageId: 'd1-p1', x: 0.76, y: 0.39, width: 0.028, height: 0.042, score: 0.86, rotation: 180, scale: 1, manual: false },
                        { id: 'match-6', pageId: 'd1-p2', x: 0.29, y: 0.31, width: 0.028, height: 0.042, score: 0.93, rotation: 0, scale: 1, manual: false },
                        { id: 'match-7', pageId: 'd1-p2', x: 0.61, y: 0.55, width: 0.028, height: 0.042, score: 0.9, rotation: 90, scale: 1, manual: false }
                    ]
                }
            });
            if (prop === 'export_symbol_count_results') return async (_sessionId, symbols = []) => ({
                status: 'success',
                data: {
                    path: 'Symbol Counter Demo Symbol Count.xlsx',
                    symbolCount: symbols.length,
                    instanceCount: symbols.reduce((total, symbol) => total + (symbol.matches || []).length, 0)
                }
            });
            if (prop === 'close_symbol_count_session') return async () => ({ status: 'success' });

            return async (...args) => {
                console.log('Mock API call: ' + prop, args);
                return { status: 'success' };
            };
        }
    })
};
window.addEventListener('load', () => {
    window.dispatchEvent(new CustomEvent('pywebviewready'));
});
