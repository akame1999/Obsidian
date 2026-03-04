"""Run this in your Windows-Patch folder to add the multi-KB queue feature."""

content = open('patcher_app.py', encoding='utf-8').read()

# 1. Add kb_queue to session state
content = content.replace(
    "    'kb_downloads': [],\n    'auto_scan_pending'",
    "    'kb_downloads': [],\n    'kb_queue': [],\n    'auto_scan_pending'"
)

# 2. Inject queue UI before the existing search bar
OLD = "    # Search bar + button\n    col_q, col_srch = st.columns([5, 1])"

NEW = """    # ── Multi-KB Queue ───────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📋 KB Queue — Download Multiple KBs At Once</div>', unsafe_allow_html=True)
    st.caption("Enter one or more KB numbers (comma or space separated), add to queue, then click Search & Download All.")

    qc1, qc2, qc3 = st.columns([4, 1, 1])
    with qc1:
        queue_input = st.text_input("Add KBs", placeholder="KB5034441, KB5005112, KB5075902",
                                    label_visibility="collapsed", key="kb_queue_input")
    with qc2:
        if st.button("➕ Add", use_container_width=True, key="kb_queue_add"):
            for k in queue_input.replace(',', ' ').split():
                k = k.strip().upper()
                if k and not k.startswith('KB'):
                    k = 'KB' + k
                if k and k not in st.session_state.kb_queue:
                    st.session_state.kb_queue.append(k)
            st.rerun()
    with qc3:
        if st.button("🗑️ Clear", use_container_width=True, key="kb_queue_clear"):
            st.session_state.kb_queue = []
            st.rerun()

    if st.session_state.kb_queue:
        st.markdown(f'<div class="info-box">📋 Queue: <strong>{", ".join(st.session_state.kb_queue)}</strong></div>', unsafe_allow_html=True)
        queue_dest = st.text_input("Download to", value=CFG.PATCH_SHARE,
                                   label_visibility="collapsed", key="kb_queue_dest",
                                   placeholder=r"\\\\server\\share\\Patches")
        if st.button(f"⬇️ Search & Download All {len(st.session_state.kb_queue)} KB(s)",
                     type="primary", use_container_width=True, key="kb_queue_run"):
            q_results = []
            q_prog = st.progress(0)
            q_status = st.empty()
            for qi, kb in enumerate(list(st.session_state.kb_queue)):
                q_status.markdown(f"🔍 Searching **{kb}** ({qi+1}/{len(st.session_state.kb_queue)})…")
                raw = catalog_search(kb)
                hits = [r for r in raw if not r.get('error') and r.get('uid') and kb in r.get('kb','').upper()]
                if not hits:
                    hits = [r for r in raw if not r.get('error') and r.get('uid')]
                if hits:
                    r = hits[0]
                    prod = r.get('products', '')
                    ext = '.cab' if 'cab' in r.get('title','').lower() else '.msu'
                    os_tag = ('_2022' if '2022' in prod else '_2019' if '2019' in prod else
                              '_2016' if '2016' in prod else '_2012R2' if '2012' in prod else
                              '_Win11' if 'windows 11' in prod.lower() else
                              '_Win10' if 'windows 10' in prod.lower() else '')
                    fname = f"{kb}{os_tag}{ext}"
                    q_status.markdown(f"⬇️ Downloading **{fname}** ({qi+1}/{len(st.session_state.kb_queue)})…")
                    res = download_kb_to_share(uid=r['uid'], filename=fname, dest_folder=queue_dest.strip())
                    ok = res.get('success', False)
                    q_results.append({'KB': kb, 'File': fname, 'Status': '✅ Done' if ok else f"❌ {res.get('error','Failed')}"})
                    st.session_state.kb_downloads.append({
                        'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'kb': kb, 'title': r.get('title','')[:60], 'filename': fname,
                        'size_mb': res.get('size_mb', 0), 'dest': queue_dest, 'status': '✅ Success' if ok else '❌ Failed'
                    })
                else:
                    q_results.append({'KB': kb, 'File': '—', 'Status': '❌ Not found in catalog'})
                q_prog.progress((qi + 1) / len(st.session_state.kb_queue))
            q_status.empty()
            ok_n = sum(1 for x in q_results if '✅' in x['Status'])
            if ok_n == len(q_results):
                st.markdown(f'<div class="ok-box">✅ All <strong>{ok_n}</strong> KBs downloaded!</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warn-box">⚠️ {ok_n}/{len(q_results)} succeeded.</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(q_results), use_container_width=True, hide_index=True)
            st.session_state.kb_queue = []
            st.session_state.patches = []

    st.divider()

    # Search bar + button
    col_q, col_srch = st.columns([5, 1])"""

if OLD in content:
    content = content.replace(OLD, NEW, 1)
    open('patcher_app.py', 'w', encoding='utf-8').write(content)
    print("SUCCESS - Queue feature added!")
else:
    print("ERROR - Could not find insertion point. Pattern not matched.")
