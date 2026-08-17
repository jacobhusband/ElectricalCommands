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
