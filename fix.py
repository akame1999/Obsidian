content = open('patcher_app.py', encoding='utf-8').read()
if 'valid_results' in content:
    print('Already fixed!')
else:
    content = content.replace(
        'st.divider()\n\n            # ',
        'valid_results = [r for r in results if r.get("uid")]\n\n            st.divider()\n\n            # ',
        1
    )
    open('patcher_app.py', 'w', encoding='utf-8').write(content)
    print('FIXED')
