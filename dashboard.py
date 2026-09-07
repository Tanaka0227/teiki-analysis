import streamlit as st
import pandas as pd
import io
from collections import defaultdict

st.set_page_config(page_title="STRAC算出ツール", page_icon="📊", layout="wide")

st.title("媒体別 STRAC算出ツール")
st.caption("ECforceの受注CSVと原価マスタをアップロードして分析")

def is_valid_product(name):
    import re
    return bool(re.search(r'\d+$', str(name).strip()))

def is_default_product(name):
    name = str(name)
    if "サービス" in name:
        return False
    if "都度" in name:
        return False
    if "【極】" in name:
        return False
    if "【一緒にお特便】" in name:
        return False
    if "アサイベリープラチナアイ" in name:
        return True
    if "きんいん" in name and "きんいん極" not in name:
        return True
    if "豆美膳" in name:
        return True
    if "歩クミン" in name:
        return True
    return False

def load_csv(files):
    all_rows = []
    for f in files:
        raw = f.read()
        for enc in ['cp932', 'utf-8', 'shift_jis']:
            try:
                text = raw.decode(enc)
                if '定期' in text:
                    break
            except (UnicodeDecodeError, LookupError):
                continue
        else:
            text = raw.decode('cp932', errors='replace')
        df = pd.read_csv(io.StringIO(text), dtype=str)
        df.columns = df.columns.str.strip()
        all_rows.append(df)
    return pd.concat(all_rows, ignore_index=True)

def load_master(f):
    df = pd.read_excel(f, dtype=str)
    df.columns = df.columns.str.strip()
    return df

def calculate(csv_df, master_df, excluded_products):
    product_map = {}
    for _, r in master_df.iterrows():
        code = str(r.get('商品コード', '')).strip()
        if not code:
            continue
        try:
            bags = int(r.get('個数', 1))
        except (ValueError, TypeError):
            bags = 1
        try:
            cost = float(r.get('原価', 0))
        except (ValueError, TypeError):
            cost = 0
        omatome = r.get('おまとめ計算', '')
        is_upsell = str(omatome).strip() == '1'
        name = str(r.get('商品名', ''))
        product_map[code] = {'bags': bags, 'is_upsell': is_upsell, 'name': name, 'cost': cost}

    teiki_to_media = {}
    teiki_n = defaultdict(lambda: defaultdict(lambda: {'codes': [], 'subtotal': 0}))
    excluded_teikis = set()

    for _, row in csv_df.iterrows():
        n_str = str(row.get('定期内受注（N番目）', '')).strip()
        teiki = str(row.get('定期受注番号', '')).strip()
        media = str(row.get('広告URLグループ名', '')).strip()
        code = str(row.get('購入商品（商品コード）', '')).strip()
        product_name = str(row.get('購入商品（商品名）', '')).strip()
        try:
            subtotal = float(row.get('小計', 0) or 0)
        except (ValueError, TypeError):
            subtotal = 0
        try:
            n = int(float(n_str))
        except (ValueError, TypeError):
            continue
        if not teiki or teiki == 'nan':
            continue
        if not is_valid_product(product_name):
            continue
        if n == 1 and product_name in excluded_products:
            excluded_teikis.add(teiki)
            continue
        if teiki in excluded_teikis:
            continue
        if n == 1 and media and media != 'nan':
            teiki_to_media[teiki] = media
        teiki_n[teiki][n]['codes'].append(code)
        teiki_n[teiki][n]['subtotal'] += subtotal

    media_stats = {}
    for teiki, n_map in teiki_n.items():
        if teiki in excluded_teikis:
            continue
        media = teiki_to_media.get(teiki)
        if not media:
            continue
        if media not in media_stats:
            media_stats[media] = {
                'q1': 0, 'q2': 0, 'upsell_count': 0, 'total_cost': 0,
                'n_counts': defaultdict(set), 'n_revenue': defaultdict(float),
                'n_cost': defaultdict(float), 'product_name': ''
            }
        ms = media_stats[media]
        for n, data in n_map.items():
            ms['n_counts'][n].add(teiki)
            ms['n_revenue'][n] += data['subtotal']
            n_cost = sum(product_map.get(c, {}).get('cost', 0) for c in data['codes'])
            ms['n_cost'][n] += n_cost
            if n == 1:
                total_bags = 0
                has_upsell = False
                for c in data['codes']:
                    prod = product_map.get(c)
                    if prod:
                        total_bags += prod['bags']
                        if prod['is_upsell']:
                            has_upsell = True
                        if not ms['product_name']:
                            ms['product_name'] = prod['name']
                    else:
                        total_bags += 1
                ms['q1'] += total_bags
                ms['q2'] += 1
                ms['total_cost'] += n_cost
                if has_upsell:
                    ms['upsell_count'] += 1

    results = []
    for media, ms in media_stats.items():
        n1 = ms['q2']
        if n1 == 0:
            continue
        upsell_rate = ms['upsell_count'] / n1
        rates = []
        for i in range(1, 13):
            cnt = len(ms['n_counts'].get(i, set()))
            rates.append(cnt / n1 if cnt > 0 else None)
        rev6 = sum(ms['n_revenue'].get(n, 0) for n in range(1, 7))
        rev12 = sum(ms['n_revenue'].get(n, 0) for n in range(1, 13))
        cost6 = sum(ms['n_cost'].get(n, 0) for n in range(1, 7))
        cost12 = sum(ms['n_cost'].get(n, 0) for n in range(1, 13))
        p_val = round((ms['n_revenue'].get(1, 0)) / n1)
        v_val = round(ms['total_cost'] / n1)
        results.append({
            '広告URLグループ': media,
            '商品名': ms['product_name'],
            'P（単価）': p_val,
            'V（原価）': v_val,
            'M（粗利）': p_val - v_val,
            'Q1（袋数）': ms['q1'],
            'Q2（件数）': n1,
            'アップセルQ': ms['upsell_count'],
            'アップセル率': upsell_rate,
            **{f'F{i}': rates[i-1] for i in range(1, 13)},
            '半年LTV（粗利）': round((rev6 - cost6) / n1),
            '一年LTV（粗利）': round((rev12 - cost12) / n1),
        })
    results.sort(key=lambda x: x['Q2（件数）'], reverse=True)
    return results, teiki_to_media

col1, col2 = st.columns(2)
with col1:
    csv_files = st.file_uploader("受注データ CSV（複数可）", type=["csv"], accept_multiple_files=True)
with col2:
    master_file = st.file_uploader("原価マスタ .xlsx", type=["xlsx", "xls"])

if csv_files and master_file:
    csv_df = load_csv(csv_files)
    master_df = load_master(master_file)

    all_products = csv_df['購入商品（商品名）'].dropna().str.strip().unique()
    valid_products = sorted([p for p in all_products if is_valid_product(p)])

    default_checked = [p for p in valid_products if is_default_product(p)]
    default_unchecked = [p for p in valid_products if not is_default_product(p)]

    tab_result, tab_products = st.tabs(["STRAC結果", "商品設定"])

    with tab_products:
        st.markdown(f"**{len(valid_products)}商品**（末尾に数字のないものは除外済み）")
        filter_text = st.text_input("商品名で絞り込み", "")

        col_a, col_b, col_c, col_d = st.columns(4)
        with col_a:
            select_all = st.button("すべて選択")
        with col_b:
            deselect_all = st.button("すべて解除")
        with col_c:
            select_filtered = st.button("絞り込み結果を選択")
        with col_d:
            deselect_filtered = st.button("絞り込み結果を解除")

        if 'selected_products' not in st.session_state:
            st.session_state.selected_products = set(default_checked)

        if select_all:
            st.session_state.selected_products = set(valid_products)
            st.rerun()
        if deselect_all:
            st.session_state.selected_products = set()
            st.rerun()

        filtered = [p for p in valid_products if not filter_text or filter_text.lower() in p.lower()]

        if select_filtered:
            st.session_state.selected_products |= set(filtered)
            st.rerun()
        if deselect_filtered:
            st.session_state.selected_products -= set(filtered)
            st.rerun()

        for p in filtered:
            checked = p in st.session_state.selected_products
            if st.checkbox(p, value=checked, key=f"prod_{p}"):
                st.session_state.selected_products.add(p)
            else:
                st.session_state.selected_products.discard(p)

        selected = st.session_state.selected_products
        excluded = set(valid_products) - selected
        st.info(f"{len(selected)}商品を選択中 / {len(excluded)}商品を除外")

    with tab_result:
        excluded = set(valid_products) - st.session_state.get('selected_products', set(default_checked))
        results, teiki_to_media = calculate(csv_df, master_df, excluded)

        c1, c2, c3 = st.columns(3)
        c1.metric("媒体数", len(results))
        c2.metric("総Q2", f"{sum(r['Q2（件数）'] for r in results):,}")
        c3.metric("除外商品数", len(excluded))

        if results:
            df_result = pd.DataFrame(results)
            pct_cols = ['アップセル率'] + [f'F{i}' for i in range(1, 13)]
            for col in pct_cols:
                df_result[col] = df_result[col].apply(lambda x: f"{x*100:.1f}%" if x is not None else "")
            for col in ['P（単価）', 'V（原価）', 'M（粗利）', 'Q1（袋数）', 'Q2（件数）', 'アップセルQ', '半年LTV（粗利）', '一年LTV（粗利）']:
                df_result[col] = df_result[col].apply(lambda x: f"{x:,}")

            st.dataframe(df_result, use_container_width=True, hide_index=True)

            tsv = df_result.to_csv(sep='\t', index=False)
            st.download_button("TSVをダウンロード", tsv, "strac_result.tsv", "text/tab-separated-values")

            st.subheader("媒体別CSV出力")
            media_list = [r['広告URLグループ'] for r in results]
            selected_media = st.selectbox("媒体を選択", media_list)
            if selected_media:
                teikis = {t for t, m in teiki_to_media.items() if m == selected_media}
                filtered_df = csv_df[csv_df['定期受注番号'].isin(teikis)]
                csv_out = filtered_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    f"{selected_media} のCSVをダウンロード（{len(filtered_df):,}行）",
                    csv_out,
                    f"{selected_media}.csv",
                    "text/csv"
                )
else:
    st.info("CSVファイルと原価マスタをアップロードしてください")
