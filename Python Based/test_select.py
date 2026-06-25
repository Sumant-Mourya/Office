from nicegui import ui
ui.select(['Option 1', 'Option 2'], value='Option 1').props('dropdown-icon=none')
ui.select(['Option 1', 'Option 2'], value='Option 1').props('dropdown-icon="img:data:image/svg+xml;utf8,<svg viewBox=\"0 0 24 24\" fill=\"%239ca3af\" xmlns=\"http://www.w3.org/2000/svg\"><path d=\"M7 10l5 5 5-5z\"/></svg>"')
ui.add_head_html('''
<style>
.q-select__dropdown-icon { font-size: 0; width: 24px; height: 24px; background: url("data:image/svg+xml;charset=utf8,%3Csvg viewBox='0 0 24 24' fill='%239ca3af' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M7 10l5 5 5-5z'/%3E%3C/svg%3E") no-repeat center center; }
</style>
''')
ui.run(port=8081, show=False)
