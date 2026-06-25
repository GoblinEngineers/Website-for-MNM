import streamlit as st
import pandas as pd
import datetime

# 1. PAGE SETUP
st.set_page_config(page_title="MTG Hub: Elite", layout="centered")

@st.cache_data
def load_data():
    file_path = "MNM Calc.xlsm"
    try:
        df = pd.read_excel(file_path, sheet_name="Data", engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        df['Date'] = pd.to_datetime(df['Date'])
        
        # Quarter Logic
        df['Quarter'] = df['Date'].dt.month.map({
            1:'Q1', 2:'Q1', 3:'Q1', 4:'Q2', 5:'Q2', 6:'Q2',
            7:'Q3', 8:'Q3', 9:'Q3', 10:'Q4', 11:'Q4', 12:'Q4'
        })
        
        # Metadata for Week grouping
        df['Year_Week'] = df['Date'].dt.strftime('%Y-W%U')
        return df
    except Exception as e:
        st.error(f"Excel Error: {e}")
        return None

def calculate_league_stats(df_view, mode="points", min_req=8):
    if df_view.empty:
        return pd.DataFrame(), pd.DataFrame()

    stats = df_view.groupby('Player').agg(
        GP=('Result', 'count'),
        Wins=('Result', lambda x: (x.str.strip().str.lower() == 'win').sum()),
        Losses=('Result', lambda x: (x.str.strip().str.lower() == 'loss').sum()),
        Draws=('Result', lambda x: (x.str.strip().str.lower() == 'draw').sum())
    ).reset_index()

    # Win% Formula: (Wins + 0.25*Draws) / GP
    stats['Win%'] = (((stats['Wins'] + (stats['Draws'] * 0.25)) / stats['GP']) * 100).round(1)
    stats['Record'] = stats.apply(lambda x: f"{int(x['Wins'])}/{int(x['Losses'])}/{int(x['Draws'])}", axis=1)

    if mode == "points":
        # Base 100 + ((W*3) + (L*-1)) * 3
        stats['Rating'] = 100 + ((stats['Wins'] * 3) + (stats['Losses'] * -1)) * 3
    else:
        # Championship System: Rank by Win% (Scaled x10)
        stats['Rating'] = (stats['Win%'] * 10).round(0)

    # Calculate Global Ranks before filtering for the "Would Be" logic
    stats = stats.sort_values('Rating', ascending=False)
    stats['Would_Be'] = range(1, len(stats) + 1)

    qualified = stats[stats['GP'] >= min_req].copy()
    bench = stats[stats['GP'] < min_req].copy()

    if not qualified.empty:
        qualified['Rank'] = range(1, len(qualified) + 1)
        qualified = qualified[['Rank', 'Player', 'Rating', 'Win%', 'Record', 'GP']]
    
    if not bench.empty:
        bench['Needed'] = min_req - bench['GP']
        bench = bench[['Would_Be', 'Player', 'GP', 'Needed']]
        
    return qualified, bench

# --- APP START ---
raw_df = load_data()

if raw_df is not None:
    # --- AUTO-DETECT DATES FOR DEFAULTS ---
    now = datetime.datetime.now()
    curr_year = now.year
    curr_q_label = f"Q{(now.month - 1) // 3 + 1}"

    st.sidebar.title("League Season")
    available_years = sorted(raw_df['Year'].unique().tolist(), reverse=True)
    try:
        def_year_idx = available_years.index(curr_year)
    except:
        def_year_idx = 0
    sel_year = st.sidebar.selectbox("Year", available_years, index=def_year_idx)

    if sel_year >= 2026:
        q_options = ["Q1", "Q2", "Q3", "Q4 (Championship)"]
        def_q_idx = next((i for i, opt in enumerate(q_options) if curr_q_label in opt), 0)
        sel_q = st.sidebar.selectbox("Quarter", q_options, index=def_q_idx)
        
        if "Q4" in sel_q:
            season_df = raw_df[raw_df['Year'] == sel_year]
            mode, min_req = "win_percent", 32
            desc = "🏆 **Championship Mode**: Ranked by Win% over all 4 quarters. (32 game min.)"
        else:
            season_df = raw_df[(raw_df['Year'] == sel_year) & (raw_df['Quarter'] == sel_q[:2])]
            mode, min_req = "points", 8
            desc = f"📈 **Points Mode ({sel_q[:2]})**: Base 100 Rating. Win +9, Loss -3 (8 game min)"
    else:
        season_df = raw_df[raw_df['Year'] == sel_year]
        mode, min_req = "points", 8
        desc = "📜 **Legacy Mode**: Historic scoring (Win +9, Loss -3). (8 game min)"

    # --- RAGAVAN CHECK ---
    if season_df.empty:
        st.markdown("<h2 style='text-align: center;'>⚔️ The MNM Stats</h2>", unsafe_allow_html=True)
        st.write("")
        st.warning(f"🐒 **The data you are looking for was stolen by Ragavan!**")
        st.info(f"There are no recorded matches for **{sel_year} {sel_q if sel_year >= 2026 else ''}** yet. Try selecting a different season in the sidebar.")
        st.stop()

    leaderboard, bench = calculate_league_stats(season_df, mode, min_req)

    st.markdown("<h2 style='text-align: center;'>⚔️ The MNM Stats</h2>", unsafe_allow_html=True)
    st.info(desc)

    tab1, tab2, tab3 = st.tabs(["🏆 Leaderboard", "👤 My Stats", "📜 Log"])

    # --- TAB 1: LEADERBOARD ---
    with tab1:
        def style_rows(row):
            gold = 'background-color: rgba(212, 175, 55, 0.35)' 
            silver = 'background-color: rgba(192, 192, 192, 0.15)'
            if row['Rank'] <= 5: return [gold] * len(row)
            if row['Rating'] > 100 or row['Win%'] > 25: return [silver] * len(row)
            return [''] * len(row)

        if not leaderboard.empty:
            st.dataframe(
                leaderboard.style.apply(style_rows, axis=1), 
                column_config={
                    "Win%": st.column_config.NumberColumn("Win%", format="%.1f%%"),
                    "Rating": st.column_config.NumberColumn("Rating", format="%d")
                },
                hide_index=True, use_container_width=True
            )

        # --- THE BENCH (Restored) ---
        if not bench.empty:
            st.divider()
            st.markdown("### ⏳ The Bench")
            st.caption(f"Unranked players who need {int(min_req)} games to qualify.")
            st.dataframe(
                bench, 
                column_config={
                    "Would_Be": "Rank?",
                    "Needed": "Games More"
                },
                hide_index=True, use_container_width=True
            )

    # --- TAB 2: MY STATS ---
    with tab2:
        player_sel = st.selectbox("Search Player", sorted(raw_df['Player'].unique()))
        p_data = raw_df[raw_df['Player'] == player_sel]
        
        # Lucky Seat (4-player only)
        lucky_data = p_data[p_data['Pod Size'] == 4]
        if not lucky_data.empty:
            seat_stats = lucky_data.groupby('Seat')['Result'].apply(lambda x: (x.str.lower() == 'win').sum()/len(x))
            
            # Check if seat_stats actually has results before getting the max
            if not seat_stats.empty:
                best_seat = int(seat_stats.idxmax())
            else:
                best_seat = "N/A"
        else:
            best_seat = "N/A"

        c1, c2, c3 = st.columns(3)
        c1.metric("Games Played", len(p_data))
        t_wins = (p_data['Result'].str.lower() == 'win').sum()
        t_draws = (p_data['Result'].str.lower() == 'draw').sum()
        c2.metric("Total Win %", f"{((t_wins + (t_draws * 0.25)) / len(p_data)) * 100:.1f}%")
        c3.metric("Lucky Seat", f"Seat {best_seat}")

        # --- NEMESIS LIST ---
        st.markdown("#### 💀 The Nemesis List")
        my_pod_ids = p_data['Pod ID'].unique()
        pod_winners = raw_df[(raw_df['Pod ID'].isin(my_pod_ids)) & 
                             (raw_df['Result'].str.lower() == 'win') & 
                             (raw_df['Player'] != player_sel)]
        
        if not pod_winners.empty:
            nemesis = pod_winners['Player'].value_counts().head(3)
            nem_cols = st.columns(len(nemesis))
            for i, (name, count) in enumerate(nemesis.items()):
                nem_cols[i].markdown(f"**{name}**")
                nem_cols[i].caption(f"{count} wins against you")
        else:
            st.write("No Nemesis found yet!")

        st.markdown("#### Historical Performance")
        def get_graph_label(row):
            if row['Year'] <= 2025: return "2025 Season"
            return f"{row['Year']} {row['Quarter']}"
        p_data['Graph_Label'] = p_data.apply(get_graph_label, axis=1)
        graph_data = p_data.groupby('Graph_Label')['Result'].apply(lambda g: ((g.str.lower() == 'win').sum() + (g.str.lower() == 'draw').sum()*0.25)/len(g)*100).reset_index()
        st.bar_chart(graph_data.set_index('Graph_Label'))

    # --- TAB 3: MATCH LOG ---
    with tab3:
        all_weeks = sorted(raw_df['Year_Week'].unique(), reverse=True)
        for i, week_label in enumerate(all_weeks[:2]):
            st.subheader("🔥 Most Recent Week" if i == 0 else "⏪ Previous Week")
            week_df = raw_df[raw_df['Year_Week'] == week_label].sort_values(['Date', 'Round'], ascending=[False, True])
            
            for rnd in sorted(week_df['Round'].unique()):
                st.markdown(f"<p style='color: gray; font-size: 11px; margin-bottom: -15px; margin-top: 10px;'>ROUND {int(rnd)}</p>", unsafe_allow_html=True)
                st.divider()
                
                round_df = week_df[week_df['Round'] == rnd]
                for pid in round_df['Pod ID'].unique():
                    pod = round_df[round_df['Pod ID'] == pid]
                    winners = pod[pod['Result'].str.lower() == 'win']['Player'].tolist()
                    st.markdown(f"""<div style="background: rgba(255,255,255,0.05); padding: 10px; border-radius: 10px; margin-bottom: 5px; border-left: 5px solid #00ffcc;">
                        <b>{pod['Date'].iloc[0].strftime('%m/%d')} Winner: {", ".join(winners) if winners else "Draw"}</b><br>
                        <small>{", ".join(pod['Player'].tolist())}</small></div>""", unsafe_allow_html=True)