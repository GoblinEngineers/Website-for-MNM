import streamlit as st
import pandas as pd
import datetime
import plotly.express as px

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

    # 1. Calculate the core stats
    stats = df_view.groupby('Player').agg(
        GP=('Result', 'count'),
        Wins=('Result', lambda x: (x.str.strip().str.lower() == 'win').sum()),
        Losses=('Result', lambda x: (x.str.strip().str.lower() == 'loss').sum()),
        Draws=('Result', lambda x: (x.str.strip().str.lower() == 'draw').sum())
    ).reset_index()

    # 2. Win% Formula: (Wins + 0.25*Draws) / GP
    stats['Win%'] = (((stats['Wins'] + (stats['Draws'] * 0.25)) / stats['GP']) * 100).round(1)
    
    # 3. Create the Record string (W/L/D)
    stats['Record'] = stats.apply(lambda x: f"{int(x['Wins'])}/{int(x['Losses'])}/{int(x['Draws'])}", axis=1)

    # 4. Calculate Rating based on mode
    if mode == "points":
        # Base 100 + ((W*3) + (L*-1)) * 3
        stats['Rating'] = 100 + ((stats['Wins'] * 3) + (stats['Losses'] * -1)) * 3
    else:
        # Championship System: Rank by Win% (Scaled x10)
        stats['Rating'] = (stats['Win%'] * 10).round(0)

    # 5. Sort by Rating to determine "Potential" rank
    stats = stats.sort_values('Rating', ascending=False)
    stats['Would_Be'] = range(1, len(stats) + 1)

    # 6. Split into Qualified and Bench
    qualified = stats[stats['GP'] >= min_req].copy()
    bench = stats[stats['GP'] < min_req].copy()

    if not qualified.empty:
        qualified['Rank'] = range(1, len(qualified) + 1)
        qualified = qualified[['Rank', 'Player', 'Rating', 'Win%', 'Record', 'GP']]
    
    if not bench.empty:
        # Show Potential Rank, Player, W/L/D Record, and Total Games
        bench = bench[['Would_Be', 'Player', 'Record', 'GP']]
        
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
        st.warning(f"🐒 **The data you are looking for was stolen by Ragavan!**")
        st.info(f"There are no recorded matches for **{sel_year}** yet.")
        st.stop()

    leaderboard, bench = calculate_league_stats(season_df, mode, min_req)

    st.markdown("<h2 style='text-align: center;'>⚔️ The MNM Stats</h2>", unsafe_allow_html=True)
    st.info(desc)

    # Define the top-level tabs
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

        # --- THE BENCH (Correctly nested inside Tab 1) ---
        if not bench.empty:
            st.divider()
            st.markdown("### ⏳ The Bench")
            st.caption(f"Unranked players (Under {int(min_req)} games)")
            st.dataframe(
                bench, 
                column_config={
                    "Would_Be": "Rank?",
                    "Record": "W/L/D",
                    "GP": "Games"
                },
                hide_index=True, 
                use_container_width=True
            )

    # --- TAB 2: MY STATS ---
    with tab2:
        player_sel = st.selectbox("Search Player", sorted(raw_df['Player'].unique()))
        p_data = raw_df[raw_df['Player'] == player_sel].copy()
        p_data = p_data.sort_values(['Date', 'Round'])
        
        # 1. Lifetime Logic
        t_wins = (p_data['Result'].str.lower() == 'win').sum()
        t_draws = (p_data['Result'].str.lower() == 'draw').sum()
        total_gp = len(p_data)
        lifetime_wr = ((t_wins + (t_draws * 0.25)) / total_gp * 100) if total_gp > 0 else 0
        
        # 2. Form Logic (Last 32 games)
        p_data['Win_Val'] = (p_data['Result'].str.lower() == 'win').astype(int) + ((p_data['Result'].str.lower() == 'draw').astype(int) * 0.25)
        recent_32 = p_data.tail(32)
        form_32_wr = (recent_32['Win_Val'].mean() * 100) if not recent_32.empty else 0

        # Metrics Display
        c1, c2, c3 = st.columns(3)
        c1.metric("Games Played", total_gp)
        c2.metric("Total WR", f"{lifetime_wr:.1f}%")
        c3.metric("Last 32 games WR", f"{form_32_wr:.1f}%")

        # Lucky Seat (4-player only)
        lucky_data = p_data[p_data['Pod Size'] == 4]
        best_seat = "N/A"
        if not lucky_data.empty:
            seat_stats = lucky_data.groupby('Seat')['Result'].apply(lambda x: (x.str.lower() == 'win').sum()/len(x))
            if not seat_stats.empty:
                best_seat = f"Seat {int(seat_stats.idxmax())}"

        st.divider()
        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown("#### 💀 Nemesis List")
            my_pod_ids = p_data['Pod ID'].unique()
            pod_winners = raw_df[(raw_df['Pod ID'].isin(my_pod_ids)) & (raw_df['Result'].str.lower() == 'win') & (raw_df['Player'] != player_sel)]
            if not pod_winners.empty:
                nemesis = pod_winners['Player'].value_counts().head(3)
                for name, count in nemesis.items():
                    st.write(f"**{name}**: {count} wins vs you")
            else:
                st.write("No Nemesis found yet!")
        with col_right:
            st.markdown("#### 🍀 Lucky Seat")
            st.write(f"Best position: **{best_seat}**")

        st.divider()
        st.markdown("#### ⚡ EMA Win Rate")
        
        WINDOW = 32
        PRIOR_WR = 0.25
        wins_list = p_data['Win_Val'].tolist()
        history = []
        for i in range(len(wins_list)):
            start = max(0, i - WINDOW + 1)
            current_window = wins_list[start : i + 1]
            real_count = len(current_window)
            fake_count = max(0, WINDOW - real_count)
            total_wins = sum(current_window) + (fake_count * PRIOR_WR)
            history.append((total_wins / WINDOW) * 100)
        
        p_data['Momentum_WR'] = history
        session_df = p_data.groupby('Date').agg({'Momentum_WR': 'last'}).reset_index()
        session_df['EMA_WR'] = session_df['Momentum_WR'].ewm(span=3, adjust=False).mean()

        p_min, p_max = session_df['EMA_WR'].min(), session_df['EMA_WR'].max()
        y_bottom = max(0, min(p_min - 8, 20))
        y_top = min(100, max(p_max + 8, 35))

        fig = px.area(
            session_df, x='Date', y='EMA_WR',
            template="plotly_dark",
            labels={'EMA_WR': 'Win Rate %', 'Date': 'Session'},
            hover_data={'Date': '|%b %d, %Y', 'EMA_WR': ':.1f'}
        )
        fig.update_traces(
            line_shape='spline', line_color='#00ffcc', 
            fillcolor='rgba(0, 255, 204, 0.1)', line_width=4,
            mode='lines+markers', marker=dict(size=7, line=dict(width=1, color='white'))
        )
        fig.update_layout(
            hovermode="x unified", height=300, margin=dict(l=0, r=0, t=10, b=0),
            yaxis=dict(range=[y_bottom, y_top], ticksuffix="%"),
            xaxis=dict(showgrid=False, tickformat="%b %d")
        )
        fig.add_hline(y=25, line_dash="dash", line_color="gray", opacity=0.4)
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

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