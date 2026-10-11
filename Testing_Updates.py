import html
import numpy as np
import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
from fractions import Fraction
from collections import Counter
from itertools import groupby

# =====================================================================
# LEAGUE SETTINGS - change scoring here, nothing else needs touching
# =====================================================================
BASE_RATING = 100            # everyone starts here
WIN_PTS = 9
LOSS_PTS = -3
DRAW_PTS = 0

# Q4 carry-in: how much Q1-Q3 results count toward the Q4 starting score
CARRY_WIN_PTS = 3            # 1/3 of WIN_PTS
CARRY_LOSS_PTS = -1          # 1/3 of LOSS_PTS
CARRY_DRAW_PTS = 0
CARRY_FLOOR_AT_ZERO = True   # a rough Q1-Q3 starts you at 0, never below

MIN_GAMES_Q = 12             # games needed to be ranked in Q1-Q3
MIN_GAMES_LEGACY = 8         # games needed to be ranked in pre-championship (legacy) years
TOP_N = 5                    # top finishers: gold rows on the board + the Contender badge
MIN_GAMES_YEAR = 32          # TOTAL games this year (Q1-Q4) needed to be ranked in the Q4 championship
CHAMPIONSHIP_START_YEAR = 2026

# Quirky badge thresholds
STREAK_STOP_LEN = 3          # Streak Stopper: beat someone riding at least this many wins in a row
PHOENIX_LOSSES = 8           # Phoenix: win right after this many losses in a row
DEAD_MISSED = 8              # Back From the Dead: Mondays missed before the comeback win
REGULAR_RUN = 8              # Regular Regular: this many league Mondays in a row
BRIDESMAID_TIMES = 2         # Bridesmaid: quarters finished 2nd on the board
REVENGE_LOSSES = 3           # Revenge Arc: losses to the same player in a row before you beat them
ROBIN_MIN_GAMES = 8          # Round Robin: who counts as a regular in a quarter
DRAW_RUN = 3                 # Stalemate Streak: draws in a row
BOARD_MIN_FOR_DARK_HORSE = 8 # Dark Horse needs at least this many ranked players on the board
IRONCLAD_RUN = 6             # Ironclad: games in a row without a loss
COLD_RUN = 10                # Cold Snap: losses in a row
HOT_HAND_RUN = 3             # Hot Hand: Clean Sweeps in a row
FIVE_ALIVE_RUN = 5           # Five-Alive: Mondays in a row with at least one win
RIVAL_BEATS = 5              # Rival Slayer: times you beat the same player
OWNER_BEATS = 8              # Owner: times you beat the same player
GLOW_POINTS = 15             # Glow Up: win-rate points gained vs. the previous quarter
WIN_MILESTONES = (10, 25, 50)  # Double Digits / Quarter Century / Half Century
PEACE_DRAWS = 20             # Peacekeeper: lifetime draws
SCARRED_LOSSES = 50          # Battle Scarred: lifetime losses
NEXT_UP_SHOW = 5               # rows shown in My Stats > Next Up before 'Show more'
IGNORED_PLAYERS = {"Unknown"}  # placeholder names that shouldn't earn badges

# Percentile brackets (by lifetime games played)
BRACKET_SIZE = 50            # games per bracket: 0-49, 50-99, 100-149, ...
BRACKET_TOP = 200            # this many games and up all share one bracket
MIN_BRACKET_PLAYERS = 5      # thinner brackets get merged into the one below

# 1. PAGE SETUP
st.set_page_config(page_title="MTG Hub: Elite", layout="centered")


@st.cache_data
def load_data():
    file_path = "MNM Calc.xlsm"
    try:
        df = pd.read_excel(file_path, sheet_name="Data", engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        df['Date'] = pd.to_datetime(df['Date'])

        if 'Year' not in df.columns:
            df['Year'] = df['Date'].dt.year

        # One cleaned result column used everywhere (fixes stray spaces / caps)
        df['Res'] = df['Result'].astype(str).str.strip().str.lower()

        # Quarter Logic
        df['Quarter'] = df['Date'].dt.month.map({
            1: 'Q1', 2: 'Q1', 3: 'Q1', 4: 'Q2', 5: 'Q2', 6: 'Q2',
            7: 'Q3', 8: 'Q3', 9: 'Q3', 10: 'Q4', 11: 'Q4', 12: 'Q4'
        })

        # Metadata for Week grouping
        df['Year_Week'] = df['Date'].dt.strftime('%Y-W%U')
        return df
    except Exception as e:
        st.error(f"Excel Error: {e}")
        return None


def tally(df):
    """Wins / losses / draws / games per player for a slice of the data."""
    d = df.assign(
        _w=(df['Res'] == 'win').astype(int),
        _l=(df['Res'] == 'loss').astype(int),
        _d=(df['Res'] == 'draw').astype(int),
    )
    return d.groupby('Player').agg(
        GP=('Res', 'count'),
        Wins=('_w', 'sum'),
        Losses=('_l', 'sum'),
        Draws=('_d', 'sum'),
    ).reset_index()


def calc_carry(q123_df):
    """Q4 carry-in: Q1-Q3 results at the carry rate, floored at 0 if enabled."""
    t = tally(q123_df)
    t['Carry'] = (t['Wins'] * CARRY_WIN_PTS
                  + t['Losses'] * CARRY_LOSS_PTS
                  + t['Draws'] * CARRY_DRAW_PTS)
    if CARRY_FLOOR_AT_ZERO:
        t['Carry'] = t['Carry'].clip(lower=0)
    return t[['Player', 'Carry']]


def _rank_and_split(stats, min_req, show_carry):
    """Win%, record string, sort (Rating > Win% > GP), then ranked vs bench."""
    gp_safe = stats['GP'].clip(lower=1)
    stats['Win%'] = (((stats['Wins'] + stats['Draws'] * 0.25) / gp_safe) * 100).round(1)
    stats['Record'] = stats.apply(
        lambda x: f"{int(x['Wins'])}/{int(x['Losses'])}/{int(x['Draws'])}", axis=1)

    stats = stats.sort_values(['Rating', 'Win%', 'GP'], ascending=False)
    stats['Would_Be'] = range(1, len(stats) + 1)

    qualified = stats[stats['GP'] >= min_req].copy()
    bench = stats[stats['GP'] < min_req].sort_values(['GP', 'Rating'], ascending=False).copy()   # most games first

    if not qualified.empty:
        qualified['Rank'] = range(1, len(qualified) + 1)
        cols = ['Rank', 'Player', 'Rating'] + (['Carry'] if show_carry else []) + ['Win%', 'Record', 'GP']
        qualified = qualified[cols]

    if not bench.empty:
        cols = ['Would_Be', 'Player'] + (['Carry'] if show_carry else []) + ['Record', 'GP']
        bench = bench[cols]

    return qualified, bench


def calculate_league_stats(df_view, min_req=MIN_GAMES_Q):
    """Q1-Q3 and legacy years: Rating = BASE + Wins*WIN + Losses*LOSS + Draws*DRAW."""
    stats = tally(df_view)
    if stats.empty:
        return pd.DataFrame(), pd.DataFrame()
    stats['Rating'] = (BASE_RATING
                       + stats['Wins'] * WIN_PTS
                       + stats['Losses'] * LOSS_PTS
                       + stats['Draws'] * DRAW_PTS).astype(int)
    return _rank_and_split(stats, min_req, show_carry=False)


def calculate_q4_stats(year_df, min_req=MIN_GAMES_YEAR):
    """
    Q4 championship.
      Rating = BASE + Carry-in (Q1-Q3 at 1/3 rate) + Q4 points
      Win% / Record / GP = whole-year numbers, and the game minimum counts the whole year.
    Everyone who played at any point this year is on the board.
    """
    stats = tally(year_df)
    if stats.empty:
        return pd.DataFrame(), pd.DataFrame()

    carry = calc_carry(year_df[year_df['Quarter'] != 'Q4'])
    q4 = tally(year_df[year_df['Quarter'] == 'Q4']).rename(
        columns={'GP': 'Q4_GP', 'Wins': 'Q4_W', 'Losses': 'Q4_L', 'Draws': 'Q4_D'})

    stats = stats.merge(carry, on='Player', how='left').merge(q4, on='Player', how='left')
    for c in ['Carry', 'Q4_GP', 'Q4_W', 'Q4_L', 'Q4_D']:
        stats[c] = stats[c].fillna(0).astype(int)

    stats['Rating'] = (BASE_RATING + stats['Carry']
                       + stats['Q4_W'] * WIN_PTS
                       + stats['Q4_L'] * LOSS_PTS
                       + stats['Q4_D'] * DRAW_PTS).astype(int)
    return _rank_and_split(stats, min_req, show_carry=True)


# =====================================================================
# PERCENTILE BRACKETS + BADGES
# =====================================================================
def bracket_percentile(df, player):
    """Win-rate percentile among players with a similar number of lifetime games."""
    t = tally(df)
    t['WR'] = (t['Wins'] + 0.25 * t['Draws']) / t['GP']
    top = BRACKET_TOP // BRACKET_SIZE
    t['RawB'] = (t['GP'] // BRACKET_SIZE).clip(upper=top)
    t['B'] = t['RawB']
    for b in range(top, 0, -1):          # thin brackets fold into the one below
        if (t['B'] == b).sum() < MIN_BRACKET_PLAYERS:
            t.loc[t['B'] == b, 'B'] = b - 1

    # a thin bottom bracket folds up into the next one instead
    groups = sorted(t['B'].unique())
    if len(groups) > 1 and (t['B'] == groups[0]).sum() < MIN_BRACKET_PLAYERS:
        t.loc[t['B'] == groups[0], 'B'] = groups[1]

    me = t[t['Player'] == player]
    if me.empty:
        return None, None, 0
    grp = t[t['B'] == me['B'].iloc[0]]
    lo_raw, hi_raw = int(grp['RawB'].min()), int(grp['RawB'].max())
    lo = lo_raw * BRACKET_SIZE
    label = f"{lo}+" if hi_raw >= top else f"{lo}-{(hi_raw + 1) * BRACKET_SIZE - 1}"

    others = grp[grp['Player'] != player]
    if others.empty:
        return label, None, len(grp)
    mine = me['WR'].iloc[0]
    pct = ((others['WR'] < mine).sum() + 0.5 * (others['WR'] == mine).sum()) / len(others) * 100
    return label, pct, len(grp)


BADGES = [  # (key, icon, name, description)
    ("belt", "🥇", "Belt Holder", "Won Q1, Q2, or Q3 and got the belt"),
    ("contender", "🥊", "Contender", f"Finished in the top {TOP_N} of a quarter"),
    ("undisputed", "👑", "Undisputed", "Won the yearly belt"),
    ("board", "📋", "On the Board", "Made the yearly scoreboard"),
    ("green", "🟢", "In the Green", f"Finished a quarter above {BASE_RATING} ({MIN_GAMES_Q}+ games)"),
    ("chump", "🧱", "Chump Block Champion", f"Went winless in a quarter ({MIN_GAMES_Q}+ games)"),
    ("diplomat", "🤝", "Diplomat", f"More draws than wins and losses combined in a quarter ({MIN_GAMES_Q}+ games)"),
    ("opening", "🃏", "Opening Hand", "Played your first game"),
    ("regular", "🎲", "Regular", "Played 50 or more games"),
    ("grinder", "⚙️", "Grinder", "Played 100 or more games"),
    ("fixture", "🪑", "Fixture", "Played 150 or more games"),
    ("legend", "🐉", "Legend", "Played 200 or more games"),
    ("hattrick", "🎩", "Hat Trick", "Won 3 games in a row"),
    ("onfire", "🔥", "On Fire", "Won 5 games in a row"),
    ("sweep", "🧹", "Clean Sweep", "Won both games on a Monday"),
    ("seats", "💺", "Seat Sampler", "Won from every seat in 4-player pods"),
    ("ironman", "🦾", "Iron Man", "Played every league week in a quarter"),
    ("rags", "📈", "Rags to Riches", f"Finished in the green right after a quarter below {BASE_RATING}"),
    ("luck", "🍀", "Beginner's Luck", "Got your first win"),
    ("b2b", "🔁", "Back-to-Back", "Won the belt in two quarters in a row"),
    ("howdid", "🤔", "How Did We Get Here?", "Played only 1 game on a Monday instead of 2"),
    ("stopper", "🛑", "Streak Stopper", f"Won a pod with a player riding {STREAK_STOP_LEN}+ wins in a row"),
    ("giant", "🗡️", "Giant Killer", "Won a pod that included the #1 ranked player"),
    ("darkhorse", "🐴", "Dark Horse", "Won a game while ranked in the bottom 3"),
    ("phoenix", "🌅", "Phoenix", f"Won right after losing {PHOENIX_LOSSES}+ games in a row"),
    ("undead", "🧟", "Back From the Dead", f"Won your first game back after missing {DEAD_MISSED}+ Mondays"),
    ("robin", "🔄", "Round Robin", f"Shared a pod with every regular ({ROBIN_MIN_GAMES}+ games) in a quarter"),
    ("bridesmaid", "💐", "Bridesmaid", f"Finished 2nd on the board in {BRIDESMAID_TIMES} different quarters"),
    ("revenge", "⚔️", "Revenge Arc", "Beat a player right after losing to them" if REVENGE_LOSSES <= 1
     else f"Beat a player after losing to them {REVENGE_LOSSES} times in a row"),
    ("dstale", "⚖️", "Double Stalemate", "Drew both games on the same Monday"),
    ("sstreak", "🧊", "Stalemate Streak", f"Drew {DRAW_RUN} games in a row"),
    ("oddtable", "🎪", "Odd Table Out", "Won in every pod size the league has played"),
    ("regreg", "📅", "Regular Regular", f"Played {REGULAR_RUN} league Mondays in a row"),
    ("rival", "🪓", "Rival Slayer", f"Beat the same player {RIVAL_BEATS} times"),
    ("owner", "🔒", "Owner", f"Beat the same player {OWNER_BEATS} times"),
    ("hothand", "✋", "Hot Hand", f"Won both games on {HOT_HAND_RUN} Mondays in a row"),
    ("deadeven", "🎯", "Dead Even", f"Finished a quarter on exactly {BASE_RATING} points"),
    ("fivealive", "🖐️", "Five-Alive", f"Won at least once on {FIVE_ALIVE_RUN} Mondays in a row"),
    ("ironclad", "🛡️", "Ironclad", f"Went {IRONCLAD_RUN} games in a row without a loss"),
    ("cold", "🥶", "Cold Snap", f"Lost {COLD_RUN} games in a row"),
    ("glow", "✨", "Glow Up", f"Raised your win rate {GLOW_POINTS}+ points from the previous quarter"),
    ("underdog", "🐕", "Underdog", "Won a pod where you were the lowest-ranked player at the table"),
    ("gauntlet", "🏰", "Gauntlet", f"Won a pod with 2 or more of the top {TOP_N}"),
    ("closer", "🔔", "Closer", "Won Round 2 on the last Monday of a quarter"),
    ("wins1", "🔟", "Double Digits", f"Won {WIN_MILESTONES[0]} games"),
    ("wins2", "🎖️", "Quarter Century", f"Won {WIN_MILESTONES[1]} games"),
    ("wins3", "💎", "Half Century", f"Won {WIN_MILESTONES[2]} games"),
    ("odds", "🎰", "Against the Odds", "Won a 4-player pod where all 3 opponents were Belt Holders"),
    ("peace", "🕊️", "Peacekeeper", f"Drew {PEACE_DRAWS} games"),
    ("scarred", "🩹", "Battle Scarred", f"Lost {SCARRED_LOSSES} games and kept coming back"),
]


def four_player_seats(df):
    """The 4 seat numbers used in 4-player pods (a stray typo can't add a 5th)."""
    seats = pd.to_numeric(df.loc[df['Pod Size'] == 4, 'Seat'], errors='coerce').dropna().astype(int)
    return set(seats.value_counts().head(4).index)


def _board_before(df, night):
    """The ranked board as it stood going into a Monday (None if there wasn't one yet)."""
    row = df[df['Date'] == night].iloc[0]
    y, q = int(row['Year']), row['Quarter']
    before = df[(df['Year'] == y) & (df['Date'] < night)]
    if before.empty:
        return None
    if y < CHAMPIONSHIP_START_YEAR:
        lb, _ = calculate_league_stats(before, MIN_GAMES_LEGACY)
    elif q != 'Q4':
        lb, _ = calculate_league_stats(before[before['Quarter'] == q], MIN_GAMES_Q)
    else:
        lb, _ = calculate_q4_stats(before, MIN_GAMES_YEAR)
    return None if lb.empty else lb


def _quirky_badges(df, add, curr_year, curr_q):
    """The odd-moment badges: streak breaks, comebacks, board upsets, rivalries."""
    d = df.sort_values(['Date', 'Round']).copy()
    d['_pk'] = list(zip(d['Date'], d['Round'], d['Pod ID']))
    nights = sorted(d['Date'].unique())
    night_idx = {n: i for i, n in enumerate(nights)}
    pods = {}
    for pk, g in d.groupby('_pk', sort=False):
        won = g.loc[g['Res'] == 'win', 'Player'].tolist()
        pods[pk] = {'players': g['Player'].tolist(), 'winner': won[0] if len(won) == 1 else None,
                    'date': g['Date'].iloc[0]}
    pod_order = sorted(pods, key=lambda k: (k[0], k[1]))
    sizes = {int(x) for x in d['Pod Size'].dropna().unique()}

    streak_before = {}
    for p, g in d.groupby('Player'):
        res, pks = g['Res'].tolist(), g['_pk'].tolist()
        wins_run = loss_run = draw_run = nl_run = 0
        for r, pk in zip(res, pks):
            streak_before[(p, pk)] = wins_run
            nl_run = nl_run + 1 if r != 'loss' else 0
            if nl_run == IRONCLAD_RUN:
                add(p, 'ironclad')
            if r == 'win':
                if loss_run >= PHOENIX_LOSSES:
                    add(p, 'phoenix')
                wins_run, loss_run, draw_run = wins_run + 1, 0, 0
            elif r == 'loss':
                wins_run, loss_run, draw_run = 0, loss_run + 1, 0
                if loss_run == COLD_RUN:
                    add(p, 'cold')
            else:
                wins_run, loss_run = 0, 0
                draw_run += 1
                if draw_run == DRAW_RUN:
                    add(p, 'sstreak')
        # Monday-level: comebacks, attendance runs, double draws
        prev_idx, run, hh_run, fa_run = None, 0, 0, 0
        for dt, night in g.groupby('Date'):
            i = night_idx[dt]
            swept = len(night) == 2 and (night['Res'] == 'win').all()
            hh_run = hh_run + 1 if swept else 0
            fa_run = fa_run + 1 if (night['Res'] == 'win').any() else 0
            if hh_run == HOT_HAND_RUN:
                add(p, 'hothand')
            if fa_run == FIVE_ALIVE_RUN:
                add(p, 'fivealive')
            if prev_idx is not None and i - prev_idx - 1 >= DEAD_MISSED and night['Res'].iloc[0] == 'win':
                add(p, 'undead')
            run = run + 1 if prev_idx is not None and i - prev_idx == 1 else 1
            if run == REGULAR_RUN:
                add(p, 'regreg')
            prev_idx = i
            if len(night) == 2 and (night['Res'] == 'draw').all():
                add(p, 'dstale')
        if sizes and {int(x) for x in g.loc[g['Res'] == 'win', 'Pod Size']} >= sizes:
            add(p, 'oddtable')
        n_w, n_d, n_l = (int((g['Res'] == r).sum()) for r in ('win', 'draw', 'loss'))
        for key, need in zip(('wins1', 'wins2', 'wins3'), WIN_MILESTONES):
            if n_w >= need:
                add(p, key)
        if n_d >= PEACE_DRAWS:
            add(p, 'peace')
        if n_l >= SCARRED_LOSSES:
            add(p, 'scarred')

    # streak breakers, and beating the same player again and again
    beats = Counter()
    for pk in pod_order:
        w = pods[pk]['winner']
        if w is not None:
            for o in pods[pk]['players']:
                if o == w:
                    continue
                if streak_before.get((o, pk), 0) >= STREAK_STOP_LEN:
                    add(w, 'stopper')
                beats[(w, o)] += 1
                if beats[(w, o)] == RIVAL_BEATS:
                    add(w, 'rival')
                if beats[(w, o)] == OWNER_BEATS:
                    add(w, 'owner')

    # upsets against the board, using the standings going into each Monday
    for night in nights:
        lb = _board_before(d, night)
        if lb is None:
            continue
        leader = lb.iloc[0]['Player']
        rank = dict(zip(lb['Player'], lb['Rank']))
        top = set(lb.head(TOP_N)['Player'])
        bottom = set(lb.tail(3)['Player']) if len(lb) >= BOARD_MIN_FOR_DARK_HORSE else set()
        for pk in pod_order:
            if pk[0] != night:
                continue
            w = pods[pk]['winner']
            if w is None:
                continue
            if w != leader and leader in pods[pk]['players']:
                add(w, 'giant')
            if w in bottom:
                add(w, 'darkhorse')
            others = [o for o in pods[pk]['players'] if o != w]
            if len(top & set(others)) >= 2:
                add(w, 'gauntlet')
            ranked_others = [o for o in others if o in rank]
            if w in rank and len(ranked_others) >= 2 and all(rank[o] < rank[w] for o in ranked_others):
                add(w, 'underdog')

    # rivalries: someone beats you, then you beat them the next time one of you wins a shared pod
    meetings = {}
    for pk in pod_order:
        w = pods[pk]['winner']
        if w is None:
            continue
        for o in pods[pk]['players']:
            if o != w:
                meetings.setdefault(frozenset((w, o)), []).append(w)
    for pair, seq in meetings.items():
        a, b = tuple(pair)
        for me, them in ((a, b), (b, a)):
            lost = 0
            for w in seq:
                if w == them:
                    lost += 1
                else:
                    if lost >= REVENGE_LOSSES:
                        add(me, 'revenge')
                    lost = 0

    # round robin: shared a pod with everyone who played that (finished) quarter
    for (y, q), qdf in d.groupby(['Year', 'Quarter']):
        if (int(y), int(q[1])) >= (curr_year, curr_q):
            continue
        last_night = qdf[qdf['Date'] == qdf['Date'].max()]
        for p in last_night.loc[(last_night['Round'] == last_night['Round'].max()) & (last_night['Res'] == 'win'), 'Player']:
            add(p, 'closer')
        gp_q = qdf.groupby('Player').size()
        roster = set(gp_q[gp_q >= ROBIN_MIN_GAMES].index)
        if len(roster) < 6:
            continue
        met = {p: set() for p in qdf['Player'].unique()}
        for _, g in qdf.groupby('_pk'):
            ps = g['Player'].tolist()
            for p in ps:
                met[p].update(ps)
        for p in roster:
            if roster <= met[p]:
                add(p, 'robin')


@st.cache_data
def compute_badges(df, curr_year, curr_q):
    """{player: {badge_key: times_earned}} for everyone. Only finished quarters/years count."""
    df = df.assign(_seat=pd.to_numeric(df['Seat'], errors='coerce'))
    earned = {p: {} for p in df['Player'].unique()}

    def add(p, key, n=1):
        earned[p][key] = earned[p].get(key, 0) + n

    # ---------- per-player: games, streaks, nights, seats ----------
    seats4 = four_player_seats(df)
    rounds_per_night = df.groupby('Date')['Round'].nunique()
    two_round_nights = set(rounds_per_night[rounds_per_night >= 2].index)

    for p, g in df.sort_values(['Date', 'Round']).groupby('Player'):
        add(p, 'opening')
        for key, n in (('regular', 50), ('grinder', 100), ('fixture', 150), ('legend', 200)):
            if len(g) >= n:
                add(p, key)
        res = g['Res'].tolist()
        if 'win' in res:
            add(p, 'luck')
        for k, run in groupby(res):                 # draws and losses break a streak
            if k == 'win':
                length = len(list(run))
                if length >= 3:
                    add(p, 'hattrick')
                if length >= 5:
                    add(p, 'onfire')
        for d, night in g.groupby('Date'):
            if d in two_round_nights:
                if len(night) == 1:
                    add(p, 'howdid')
                elif len(night) == 2 and (night['Res'] == 'win').all():
                    add(p, 'sweep')
        if len(seats4) >= 2:
            won = set(g.loc[(g['Pod Size'] == 4) & (g['Res'] == 'win'), '_seat'].dropna().astype(int))
            if seats4 <= won:
                add(p, 'seats')

    # ---------- quarters and years (finished ones only, official from CHAMPIONSHIP_START_YEAR) ----------
    winners, qkeys, qrating, seconds, qwr = {}, [], {}, {}, {}
    for y in sorted(int(v) for v in df['Year'].unique()):
        ydf = df[df['Year'] == y]
        if y < CHAMPIONSHIP_START_YEAR:              # legacy years: one board for the whole year
            if y < curr_year:
                lb, _ = calculate_league_stats(ydf, MIN_GAMES_LEGACY)
                if not lb.empty:
                    add(lb.iloc[0]['Player'], 'undisputed')
                    for p in lb['Player']:
                        add(p, 'board')
            continue

        for q in (1, 2, 3, 4):
            qdf = ydf[ydf['Quarter'] == f'Q{q}']
            if qdf.empty or (y, q) >= (curr_year, curr_q):
                continue

            t = tally(qdf)
            t = t[t['GP'] >= MIN_GAMES_Q]
            rating = (BASE_RATING + t['Wins'] * WIN_PTS + t['Losses'] * LOSS_PTS
                      + t['Draws'] * DRAW_PTS)
            for (_, r), rt in zip(t.iterrows(), rating):
                if rt > BASE_RATING:
                    add(r['Player'], 'green')
                if rt == BASE_RATING:
                    add(r['Player'], 'deadeven')
                if r['Wins'] == 0:
                    add(r['Player'], 'chump')
                if r['Draws'] > r['Wins'] + r['Losses']:
                    add(r['Player'], 'diplomat')
            qrating[(y, q)] = dict(zip(t['Player'], rating))
            qwr[(y, q)] = dict(zip(t['Player'], (t['Wins'] + 0.25 * t['Draws']) / t['GP'] * 100))

            all_weeks = set(qdf['Year_Week'])
            if len(all_weeks) >= 4:
                for p, weeks in qdf.groupby('Player')['Year_Week']:
                    if set(weeks) == all_weeks:
                        add(p, 'ironman')

            if q < 4:
                lb, _ = calculate_league_stats(qdf, MIN_GAMES_Q)
            else:                                    # Q4 = the yearly championship board
                lb, _ = calculate_q4_stats(ydf, MIN_GAMES_YEAR)
            if not lb.empty:
                for p in lb[lb['Rank'] <= TOP_N]['Player']:
                    add(p, 'contender')
                winners[(y, q)] = lb.iloc[0]['Player']
                if len(lb) > 1:
                    seconds[(y, q)] = lb.iloc[1]['Player']
                if q < 4:
                    add(lb.iloc[0]['Player'], 'belt')
                else:
                    add(lb.iloc[0]['Player'], 'undisputed')
                    for p in lb['Player']:
                        add(p, 'board')
            qkeys.append((y, q))

    for k1, k2 in zip(qkeys, qkeys[1:]):              # back-to-back quarters
        for p, v in qwr[k2].items():
            if p in qwr[k1] and v - qwr[k1][p] >= GLOW_POINTS:
                add(p, 'glow')
        for p, rt in qrating[k2].items():
            if rt > BASE_RATING and qrating[k1].get(p, BASE_RATING) < BASE_RATING:
                add(p, 'rags')
        if winners.get(k1) is not None and winners.get(k1) == winners.get(k2):
            add(winners[k1], 'b2b')

    for p, n in Counter(seconds.values()).items():     # always the bridesmaid
        if n >= BRIDESMAID_TIMES:
            add(p, 'bridesmaid')

    _quirky_badges(df, add, curr_year, curr_q)

    # Against the Odds: won a pod where all 3 opponents already held a belt (Q1-Q3) at the time
    belt_since = {}
    for (y, q), p in winners.items():
        if q < 4:
            last = df[(df['Year'] == y) & (df['Quarter'] == f'Q{q}')]['Date'].max()
            belt_since[p] = min(belt_since.get(p, last), last)
    for _, g in df.groupby(['Date', 'Round', 'Pod ID']):
        won = g.loc[g['Res'] == 'win', 'Player'].tolist()
        if len(g) != 4 or len(won) != 1:
            continue
        others = [o for o in g['Player'] if o != won[0]]
        if all(o in belt_since and belt_since[o] < g['Date'].iloc[0] for o in others):
            add(won[0], 'odds')

    earned = {p: v for p, v in earned.items() if p not in IGNORED_PLAYERS}
    return earned, {'winners': winners, 'qkeys': qkeys, 'qrating': qrating}


def _n(n, word):
    return f"{n} {word}" + ("" if n == 1 else "s")


def badge_progress(df, player, mine, extras, curr_year, curr_q):
    """Short 'how close am I' notes, only for badges the player hasn't earned yet."""
    prog = {}
    g = df[df['Player'] == player].sort_values(['Date', 'Round'])
    gp = len(g)

    if 'luck' not in mine:
        prog['luck'] = "Win a game to earn it"

    # games milestones
    for key, n in (('regular', 50), ('grinder', 100), ('fixture', 150), ('legend', 200)):
        if key not in mine and gp < n:
            prog[key] = f"{_n(n - gp, 'game')} to go"

    # win streaks (current run)
    cur = 0
    for r in reversed(g['Res'].tolist()):
        if r != 'win':
            break
        cur += 1
    for key, n in (('hattrick', 3), ('onfire', 5)):
        if key not in mine:
            prog[key] = f"Current streak {cur}/{n}"

    # seat sampler
    if 'seats' not in mine:
        seats4 = sorted(four_player_seats(df))
        won = set(pd.to_numeric(g.loc[(g['Pod Size'] == 4) & (g['Res'] == 'win'), 'Seat'],
                                errors='coerce').dropna().astype(int))
        need = [x for x in seats4 if x not in won]
        if len(seats4) >= 2 and need:
            prog['seats'] = "Still need a win from " + ", ".join(f"Seat {x}" for x in need)

    # current-year / current-quarter badges
    if curr_year >= CHAMPIONSHIP_START_YEAR:
        ydf = df[df['Year'] == curr_year]
        qdf = ydf[ydf['Quarter'] == f'Q{curr_q}']

        if 'board' not in mine:
            yg = int((g['Year'] == curr_year).sum())
            prog['board'] = (f"{_n(MIN_GAMES_YEAR - yg, 'game')} to go this year to be ranked"
                             if yg < MIN_GAMES_YEAR else "Minimum met, awarded when the year ends")

        # belt (Q1-Q3) / yearly belt (Q4): where you stand right now
        if curr_q < 4:
            key, need_gp = 'belt', MIN_GAMES_Q
            lb, bench = calculate_league_stats(qdf, MIN_GAMES_Q)
        else:
            key, need_gp = 'undisputed', MIN_GAMES_YEAR
            lb, bench = calculate_q4_stats(ydf, MIN_GAMES_YEAR)
        row = lb[lb['Player'] == player] if not lb.empty else lb
        brow = bench[bench['Player'] == player] if not bench.empty else bench

        if 'contender' not in mine:
            if not row.empty:
                rank = int(row['Rank'].iloc[0])
                if rank <= TOP_N:
                    prog['contender'] = f"You're #{rank} right now"
                else:
                    gap = int(lb['Rating'].iloc[TOP_N - 1] - row['Rating'].iloc[0])
                    prog['contender'] = f"Currently #{rank}, {_n(gap, 'pt')} behind #{TOP_N}"
            elif not brow.empty:
                prog['contender'] = f"{_n(need_gp - int(brow['GP'].iloc[0]), 'more game')} to be ranked"

        if key not in mine:
            if not row.empty:
                rank = int(row['Rank'].iloc[0])
                behind = int(lb['Rating'].iloc[0] - row['Rating'].iloc[0])
                prog[key] = "You're #1 right now" if rank == 1 else f"Currently #{rank}, {_n(behind, 'pt')} behind first"
            elif not brow.empty:
                prog[key] = f"{_n(need_gp - int(brow['GP'].iloc[0]), 'more game')} to be ranked"

        # in the green: this quarter so far
        if 'green' not in mine:
            t = tally(qdf)
            r = t[t['Player'] == player]
            gp_q = int(r['GP'].iloc[0]) if not r.empty else 0
            rating = BASE_RATING + (int(r['Wins'].iloc[0]) * WIN_PTS + int(r['Losses'].iloc[0]) * LOSS_PTS
                                    + int(r['Draws'].iloc[0]) * DRAW_PTS if not r.empty else 0)
            bits = [f"Rating {rating}"]
            if rating <= BASE_RATING:
                bits.append(f"{_n(-(-(BASE_RATING + 1 - rating) // WIN_PTS), 'win')} to pass {BASE_RATING}")
            if gp_q < MIN_GAMES_Q:
                bits.append(f"{_n(MIN_GAMES_Q - gp_q, 'game')} to qualify")
            prog['green'] = ", ".join(bits)

        # iron man: attendance so far this quarter
        if 'ironman' not in mine:
            weeks_all = set(qdf['Year_Week'])
            weeks_me = set(qdf.loc[qdf['Player'] == player, 'Year_Week'])
            if weeks_all:
                missed = len(weeks_all - weeks_me)
                prog['ironman'] = (f"Missed {_n(missed, 'week')} this quarter, try next quarter" if missed
                                   else f"Perfect so far: {len(weeks_all)}/{len(weeks_all)} weeks")

        # rags to riches / back-to-back depend on how last quarter went
        prev_key = (curr_year, curr_q - 1) if curr_q > 1 else (curr_year - 1, 4)
        if extras['qkeys'] and extras['qkeys'][-1] == prev_key:
            prev_rating = extras['qrating'][prev_key].get(player)
            if 'rags' not in mine and prev_rating is not None and prev_rating < BASE_RATING:
                prog['rags'] = f"Last quarter: {prev_rating}. Finish above {BASE_RATING}"
            if 'b2b' not in mine and extras['winners'].get(prev_key) == player:
                prog['b2b'] = "You won last quarter, win this one too"

    return prog


def next_up(df, player, mine, extras, curr_year, curr_q):
    """Badges the player is working toward, closest first: [{'key','frac','note','tag'}]."""
    items = []

    def put(key, cur, target, note, frac=None, tag=""):
        if key in mine:
            return
        frac = min(cur / target, 0.99) if frac is None else frac
        if frac > 0:
            items.append({'key': key, 'frac': frac, 'note': note, 'tag': tag})

    def ladder(rungs, cur, note_fn):                 # only the next unearned rung is shown
        for key, target in rungs:
            if key not in mine:
                put(key, cur, target, note_fn(target, cur))
                break

    g = df[df['Player'] == player].sort_values(['Date', 'Round'])
    res = g['Res'].tolist()
    gp = len(g)
    n_w, n_d, n_l = (res.count(r) for r in ('win', 'draw', 'loss'))

    def tail_run(ok):
        k = 0
        for r in reversed(res):
            if not ok(r):
                break
            k += 1
        return k

    ladder((('regular', 50), ('grinder', 100), ('fixture', 150), ('legend', 200)), gp,
           lambda t, c: f"{_n(t - c, 'game')} to go ({c}/{t})")
    ladder((('wins1', WIN_MILESTONES[0]), ('wins2', WIN_MILESTONES[1]), ('wins3', WIN_MILESTONES[2])), n_w,
           lambda t, c: f"{_n(t - c, 'win')} to go ({c}/{t})")
    ladder((('hattrick', 3), ('onfire', 5)), tail_run(lambda r: r == 'win'),
           lambda t, c: f"Current win streak {c}/{t}")
    put('ironclad', tail_run(lambda r: r != 'loss'), IRONCLAD_RUN,
        f"{tail_run(lambda r: r != 'loss')}/{IRONCLAD_RUN} games in a row without a loss")
    put('sstreak', tail_run(lambda r: r == 'draw'), DRAW_RUN,
        f"{tail_run(lambda r: r == 'draw')}/{DRAW_RUN} draws in a row")
    put('peace', n_d, PEACE_DRAWS, f"{_n(PEACE_DRAWS - n_d, 'more draw')} ({n_d}/{PEACE_DRAWS})")
    put('scarred', n_l, SCARRED_LOSSES, f"{SCARRED_LOSSES - n_l} more {'loss' if SCARRED_LOSSES - n_l == 1 else 'losses'} ({n_l}/{SCARRED_LOSSES})")

    # Monday-level runs (current run counting back from the latest league Monday)
    nights = sorted(df['Date'].unique())
    by_night = {dt: n for dt, n in g.groupby('Date')}

    def night_run(ok):
        k = 0
        for dt in reversed(nights):
            n = by_night.get(dt)
            if n is None or not ok(n):
                break
            k += 1
        return k

    for key, target, ok, label in (
            ('regreg', REGULAR_RUN, lambda n: True, "Mondays in a row"),
            ('fivealive', FIVE_ALIVE_RUN, lambda n: (n['Res'] == 'win').any(), "Mondays in a row with a win"),
            ('hothand', HOT_HAND_RUN, lambda n: len(n) == 2 and (n['Res'] == 'win').all(), "Mondays in a row with a sweep")):
        c = night_run(ok)
        put(key, c, target, f"{c}/{target} {label}")

    # seat sampler
    if 'seats' not in mine:
        seats4 = sorted(four_player_seats(df))
        won = set(pd.to_numeric(g.loc[(g['Pod Size'] == 4) & (g['Res'] == 'win'), 'Seat'],
                                errors='coerce').dropna().astype(int))
        need = [x for x in seats4 if x not in won]
        if len(seats4) >= 2 and need:
            put('seats', len(seats4) - len(need), len(seats4),
                "Still need a win from " + ", ".join(f"Seat {x}" for x in need))

    # beating the same player again and again
    if 'rival' not in mine or 'owner' not in mine:
        beats = Counter()
        d = df.assign(_pk=list(zip(df['Date'], df['Round'], df['Pod ID'])))
        mine_pods = d.loc[(d['Player'] == player) & (d['Res'] == 'win'), '_pk']
        for pk, pod in d[d['_pk'].isin(set(mine_pods))].groupby('_pk'):
            if (pod['Res'] == 'win').sum() == 1:
                beats.update(o for o in pod['Player'] if o != player)
        if beats:
            foe, c = beats.most_common(1)[0]
            if c < OWNER_BEATS:
                key, target = ('rival', RIVAL_BEATS) if c < RIVAL_BEATS else ('owner', OWNER_BEATS)
                put(key, c, target, f"Beat {foe} {_n(target - c, 'more time')} ({c}/{target})")

    # this year / this quarter
    if curr_year >= CHAMPIONSHIP_START_YEAR:
        ydf = df[df['Year'] == curr_year]
        qdf = ydf[ydf['Quarter'] == f'Q{curr_q}']
        tag = f"Q{curr_q} so far"
        yg = int((g['Year'] == curr_year).sum())
        if yg < MIN_GAMES_YEAR:
            put('board', yg, MIN_GAMES_YEAR, f"{_n(MIN_GAMES_YEAR - yg, 'game')} to go this year to be ranked", tag="This year")

        if curr_q < 4:
            bkey, need_gp = 'belt', MIN_GAMES_Q
            lb, bench = calculate_league_stats(qdf, MIN_GAMES_Q)
        else:
            bkey, need_gp = 'undisputed', MIN_GAMES_YEAR
            lb, bench = calculate_q4_stats(ydf, MIN_GAMES_YEAR)
        row = lb[lb['Player'] == player] if not lb.empty else lb
        brow = bench[bench['Player'] == player] if not bench.empty else bench
        if not row.empty:
            rank, rt = int(row['Rank'].iloc[0]), row['Rating'].iloc[0]
            if rank <= TOP_N:
                put('contender', 0, 1, f"You're #{rank} right now. Hold it to the end of the quarter", 0.9, tag)
            else:
                gap = int(lb['Rating'].iloc[TOP_N - 1] - rt)
                put('contender', 0, 1, f"Currently #{rank}, {_n(gap, 'pt')} behind #{TOP_N}",
                    0.1 + 0.6 * max(0.0, 1 - gap / 60), tag)
            if rank == 1:
                put(bkey, 0, 1, "You're #1 right now. Hold it to the end", 0.9, tag)
            else:
                behind = int(lb['Rating'].iloc[0] - rt)
                put(bkey, 0, 1, f"Currently #{rank}, {_n(behind, 'pt')} behind first",
                    0.1 + 0.5 * max(0.0, 1 - behind / 80), tag)
        elif not brow.empty:
            c = int(brow['GP'].iloc[0])
            put('contender', c, need_gp, f"{_n(need_gp - c, 'more game')} to be ranked", 0.4 * min(c / need_gp, 1), tag)

        t = tally(qdf)
        r = t[t['Player'] == player]
        if not r.empty:
            gp_q = int(r['GP'].iloc[0])
            rating = BASE_RATING + int(r['Wins'].iloc[0]) * WIN_PTS + int(r['Losses'].iloc[0]) * LOSS_PTS \
                + int(r['Draws'].iloc[0]) * DRAW_PTS
            if rating > BASE_RATING and gp_q >= MIN_GAMES_Q:
                put('green', 0, 1, f"At {rating} right now. Stay above {BASE_RATING} to the end", 0.9, tag)
            else:
                bits = [f"Rating {rating}"]
                if rating <= BASE_RATING:
                    bits.append(f"{_n(-(-(BASE_RATING + 1 - rating) // WIN_PTS), 'win')} to pass {BASE_RATING}")
                if gp_q < MIN_GAMES_Q:
                    bits.append(f"{_n(MIN_GAMES_Q - gp_q, 'game')} to qualify")
                put('green', 0, 1, ", ".join(bits),
                    0.1 + 0.5 * min(gp_q / MIN_GAMES_Q, 1) * (1 if rating > BASE_RATING else 0.6), tag)

        weeks_all = set(qdf['Year_Week'])
        weeks_me = set(qdf.loc[qdf['Player'] == player, 'Year_Week'])
        if len(weeks_all) >= 1 and weeks_me >= weeks_all:
            put('ironman', 0, 1, f"Perfect so far: {len(weeks_all)}/{len(weeks_all)} weeks", 0.85, tag)

        prev_key = (curr_year, curr_q - 1) if curr_q > 1 else (curr_year - 1, 4)
        if extras['qkeys'] and extras['qkeys'][-1] == prev_key:
            prev_rating = extras['qrating'][prev_key].get(player)
            if prev_rating is not None and prev_rating < BASE_RATING:
                put('rags', 0, 1, f"Last quarter: {prev_rating}. Finish above {BASE_RATING}", 0.5, tag)
            if extras['winners'].get(prev_key) == player:
                put('b2b', 0, 1, "You won last quarter, win this one too", 0.5, tag)

    items.sort(key=lambda i: -i['frac'])
    return items


# =====================================================================
# STORE STATS (league-wide numbers)
# =====================================================================
POD_KEY = ['Date', 'Round', 'Pod ID']
STORE_SIMS = 2000            # luck-only simulations used for the "same winners" tests
MIN_PODS_BREAKDOWN = 20      # decided pods a player needs to appear in the beyond-luck chart


def _pfmt(p):
    return "<0.001" if p < 0.001 else f"{p:.3f}"


def _odds(p):
    """Turn a p-value into 'about 1 in N' wording."""
    if p < 0.001:
        return "less than 1 time in 1,000"
    if p < 0.5:
        return f"about 1 time in {round(1 / p):,}"
    return "in most luck-only leagues"


@st.cache_data
def pod_table(df):
    """One row per pod: size, number of winners, winner, players."""
    d = df.assign(_w=(df['Res'] == 'win').astype(int))
    pods = d.groupby(POD_KEY, sort=False).agg(
        size=('Res', 'size'), wins=('_w', 'sum'), Year=('Year', 'first'),
        Quarter=('Quarter', 'first'), players=('Player', list)).reset_index()
    win = d[d['_w'] == 1].groupby(POD_KEY)['Player'].first().rename('winner').reset_index()
    return pods.merge(win, on=POD_KEY, how='left')


def _topk_share(counts, k):
    top = -np.sort(-counts, axis=-1)[..., :k]
    return top.sum(-1) / counts.sum(-1)


def simulate_wins(dec, sims, seed=7):
    """Luck-only league: every decided pod's winner is picked at random from its players."""
    players = sorted({p for lst in dec['players'] for p in lst})
    idx = {p: i for i, p in enumerate(players)}
    pad = np.zeros((len(dec), int(dec['size'].max())), dtype=int)
    for r, lst in enumerate(dec['players']):
        pad[r, :len(lst)] = [idx[p] for p in lst]
    rng = np.random.default_rng(seed)
    pick = (rng.random((sims, len(dec))) * dec['size'].to_numpy()).astype(int)
    won = pad[np.arange(len(dec)), pick]
    flat = (np.arange(sims)[:, None] * len(players) + won).ravel()
    counts = np.bincount(flat, minlength=sims * len(players)).reshape(sims, len(players))
    return players, idx, counts


@st.cache_data
def seat_analysis(df, size):
    """Win% by seat for pods of one size, plus a luck test for a seat advantage."""
    d = df.assign(_w=(df['Res'] == 'win').astype(int),
                  _seat=pd.to_numeric(df['Seat'], errors='coerce'))
    g = d.groupby(POD_KEY, sort=False)
    d['_n'] = g['Res'].transform('size')
    d['_pw'] = g['_w'].transform('sum')
    s = d[(d['_n'] == size) & (d['_pw'] <= 1)]
    pod_wins = s.groupby(POD_KEY)['_pw'].first()
    decisive, draws = int((pod_wins == 1).sum()), int((pod_wins == 0).sum())
    if decisive < 20:
        return None
    fair = (1 - draws / (decisive + draws)) / size * 100
    s = s[s['_seat'].between(1, size)]
    seats = list(range(1, size + 1))
    t = s.groupby('_seat').agg(Games=('Res', 'size'), Wins=('_w', 'sum')).reindex(seats).fillna(0).astype(int)
    t['Win%'] = (t['Wins'] / t['Games'].clip(lower=1) * 100).round(1)
    t['Fair%'] = round(fair, 1)
    t['Edge'] = (t['Win%'] - t['Fair%']).round(1)
    t = t.reset_index().rename(columns={'index': 'Seat', '_seat': 'Seat'})
    obs = s[(s['_pw'] == 1) & (s['_w'] == 1)].groupby('_seat').size().reindex(seats).fillna(0).to_numpy()
    exp = obs.sum() / size
    stat = ((obs - exp) ** 2 / exp).sum()
    sim = np.random.default_rng(3).multinomial(int(obs.sum()), [1 / size] * size, size=5000)
    p = float((((sim - exp) ** 2 / exp).sum(1) >= stat).mean())
    return t, p, decisive + draws


@st.cache_data
def winner_concentration(df):
    """Do the same people keep winning? Compare real wins to a luck-only league."""
    pods = pod_table(df)
    dec = pods[pods['wins'] == 1].reset_index(drop=True)
    if len(dec) < 20:
        return None
    players, idx, sims = simulate_wins(dec, STORE_SIMS)
    n = len(players)
    obs, exp, var, played = (np.zeros(n) for _ in range(4))
    for lst, size, w in zip(dec['players'], dec['size'], dec['winner']):
        q = 1 / size
        for p in lst:
            i = idx[p]
            exp[i] += q
            var[i] += q * (1 - q)
            played[i] += 1
        obs[idx[w]] += 1
    top5, sim_top5 = _topk_share(obs, 5), _topk_share(sims, 5)
    safe = np.where(exp > 0, exp, 1)
    chi_obs, chi_sims = (((obs - exp) ** 2) / safe).sum(), (((sims - exp) ** 2) / safe).sum(1)
    best = np.argsort(-obs)[:5]
    table = pd.DataFrame({
        'Player': players, 'Pods': played.astype(int), 'Wins': obs.astype(int),
        'Expected': exp.round(1), 'Above luck': (obs - exp).round(1),
        'z': ((obs - exp) / np.sqrt(np.where(var > 0, var, 1))).round(2)})
    return dict(table=table, pods=len(dec), top5=float(top5), luck_top5=float(sim_top5.mean()),
                luck_lo=float(np.percentile(sim_top5, 2.5)), luck_hi=float(np.percentile(sim_top5, 97.5)),
                p_top5=float((sim_top5 >= top5).mean()), p_chi=float((chi_sims >= chi_obs).mean()),
                games_share=float(played[best].sum() / played.sum()))


@st.cache_data
def period_table(df):
    """Per quarter: did a few people take more of the wins than dice would give?"""
    dec = pod_table(df)
    dec = dec[dec['wins'] == 1]
    rows = []
    for (y, q), grp in dec.groupby(['Year', 'Quarter']):
        if len(grp) < 10:
            continue
        grp = grp.reset_index(drop=True)
        _, _, sims = simulate_wins(grp, 1000)
        counts = grp['winner'].value_counts()
        obs3 = counts.head(3).sum() / len(grp)
        sim3 = _topk_share(sims, 3)
        p3 = float((sim3 >= obs3).mean())
        verdict = ("Same few people winning" if p3 < 0.05 else
                   "Leaning that way" if p3 < 0.20 else "Normal luck")
        rows.append({'Period': f"{int(y)} {q}", 'Games with a winner': len(grp),
                     'Different winners': len(counts),
                     'Different winners if random': round(float((sims > 0).sum(1).mean()), 1),
                     'Top-3 share of wins %': round(obs3 * 100, 1),
                     'Top-3 share if random %': round(float(sim3.mean() * 100), 1),
                     'Verdict': verdict})
    return pd.DataFrame(rows).sort_values('Period', ascending=False) if rows else pd.DataFrame()


# =====================================================================
# PHONE-FRIENDLY DISPLAY HELPERS
# =====================================================================
CHART_CONFIG = {'displayModeBar': False, 'scrollZoom': False}

MOBILE_CSS = """<style>
@media (max-width: 640px) {
  .block-container, [data-testid="stMainBlockContainer"] { padding: 1.2rem 0.7rem 3rem 0.7rem !important; }
  button[data-baseweb="tab"] { padding-left: 0.55rem !important; padding-right: 0.55rem !important; }
  button[data-baseweb="tab"] p { font-size: 0.85rem !important; }
  h2 { font-size: 1.5rem !important; }
  h4 { font-size: 1.05rem !important; }
  table.mnm-table { font-size: 12px !important; }
  .mnm-table th, .mnm-table td { padding: 5px 3px !important; }
}
.mnm-metrics { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 8px; margin: 6px 0 10px 0; }
.mnm-metric { background: rgba(128,128,128,0.14); border-radius: 14px; padding: 8px 10px; }
.mnm-metric .l { font-size: 12px; opacity: 0.7; line-height: 1.25; }
.mnm-metric .v { font-size: 22px; font-weight: 600; line-height: 1.35; }
.mnm-metric .s { font-size: 11px; opacity: 0.6; line-height: 1.25; }
.mnm-wrap { overflow-x: auto; margin-bottom: 8px; padding: 2px 6px 4px 6px; border-radius: 14px; border: 1px solid rgba(128,128,128,0.25); background: rgba(128,128,128,0.07); }
div[data-testid="stPlotlyChart"] { border-radius: 14px; overflow: hidden; background: rgba(128,128,128,0.07); border: 1px solid rgba(128,128,128,0.25); padding: 6px 4px 2px 4px; }
table.mnm-table { width: 100%; border-collapse: separate; border-spacing: 0 3px; font-size: 13px; }
.mnm-table th { text-align: right; font-weight: 600; opacity: 0.7; padding: 6px 4px; line-height: 1.15; vertical-align: bottom; }
.mnm-table td { text-align: right; padding: 6px 4px; }
.mnm-table tbody tr:nth-child(odd) td { background: var(--bg, transparent); }
.mnm-table tbody tr:nth-child(even) td { background: var(--bg, rgba(128,128,128,0.08)); }
.mnm-table td:first-child { border-radius: 10px 0 0 10px; }
.mnm-table td:last-child { border-radius: 0 10px 10px 0; }
.mnm-table th.l, .mnm-table td.l { text-align: left; word-break: break-word; }
.mnm-legend { font-size: 12.5px; margin: 2px 0 8px 0; line-height: 1.9; }
.mnm-chip { display: inline-block; padding: 1px 9px; border-radius: 999px; font-weight: 600; margin-right: 2px; }
</style>"""


def _cell(v):
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v)


def _int(v):
    return f"{int(v)}"


def _pct(v):
    return f"{v:.1f}%"


def _signed(v):
    return f"{v:+.1f}"


def metric_grid(items):
    """Stat cards that wrap two-per-row on a phone and never cut their labels off."""
    cards = ""
    for it in items:
        sub = it[2] if len(it) > 2 else ""
        cards += (f'<div class="mnm-metric"><div class="l">{html.escape(str(it[0]))}</div>'
                  f'<div class="v">{html.escape(str(it[1]))}</div>'
                  + (f'<div class="s">{html.escape(str(sub))}</div>' if sub else "") + '</div>')
    st.markdown(f'<div class="mnm-metrics">{cards}</div>', unsafe_allow_html=True)


def week_moves(now_lb, prev_lb):
    """{player: (kind, steps)} comparing each ranked player's place now vs. the board before the latest week."""
    prev = dict(zip(prev_lb['Player'], prev_lb['Rank'])) if prev_lb is not None and not prev_lb.empty else {}
    out = {}
    for p, r in zip(now_lb['Player'], now_lb['Rank']):
        if p not in prev:
            out[p] = ('new', 0)
        else:
            d = int(prev[p]) - int(r)
            out[p] = ('up' if d > 0 else 'down' if d < 0 else 'same', abs(d))
    return out


def move_badge(kind, steps):
    if kind == 'up':
        return f'<span style="color:#4ade80; font-weight:600;">▲{steps}</span>'
    if kind == 'down':
        return f'<span style="color:#ff6b6b; font-weight:600;">▼{steps}</span>'
    if kind == 'new':
        return '<span style="color:#f5b942; font-size:10px; font-weight:700;">NEW</span>'
    return '<span style="opacity:0.4;">–</span>'


def html_table(df, headers=None, fmt=None, row_bg=None, left=None, html_cols=()):
    """A table that fits the screen: headers wrap, columns share the width, nothing is hidden off to the side."""
    headers, fmt = headers or {}, fmt or {}
    cols = list(df.columns)
    if left is None:
        left = [c for c in cols if df[c].dtype == object]
    head = "".join(f'<th class="{"l" if c in left else ""}">{html.escape(str(headers.get(c, c)))}</th>' for c in cols)
    body = []
    for tup in df.itertuples(index=False, name=None):
        row = dict(zip(cols, tup))
        cells = "".join(
            f'<td class="{"l" if c in left else ""}">'
            + (str(row[c]) if c in html_cols else html.escape(fmt[c](row[c]) if c in fmt else _cell(row[c])))
            + '</td>'
            for c in cols)
        bg = row_bg(row) if row_bg else ""
        body.append(f'<tr style="{"--bg:" + bg + ";" if bg else ""}">{cells}</tr>')
    st.markdown(f'<div class="mnm-wrap"><table class="mnm-table"><thead><tr>{head}</tr></thead>'
                f'<tbody>{"".join(body)}</tbody></table></div>', unsafe_allow_html=True)


def phone_fig(fig, height):
    """Charts that can't be zoomed or dragged by accident (which would hide data) and keep their labels."""
    fig.update_layout(height=height, dragmode=False, margin=dict(l=4, r=10, t=10, b=4), font=dict(size=12))
    fig.update_xaxes(fixedrange=True, automargin=True)
    fig.update_yaxes(fixedrange=True, automargin=True)
    return fig


def quarter_cards(pt):
    """One small card per quarter: real numbers next to what dice would give."""
    colors = {"Same few people winning": "#ff6b6b", "Leaning that way": "#f5b942", "Normal luck": "#00ffcc"}
    cards = ""
    for _, r in pt.iterrows():
        c = colors.get(r['Verdict'], "#6b7280")
        cards += (
            f'<div class="mnm-metric" style="border-left:5px solid {c};">'
            f'<div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:4px;">'
            f'<b>{html.escape(str(r["Period"]))}</b><span style="color:{c}; font-size:12px;">{html.escape(r["Verdict"])}</span></div>'
            f'<table class="mnm-table" style="margin-top:4px;"><thead><tr><th class="l"></th><th>Real</th><th>Dice</th></tr></thead>'
            f'<tbody><tr><td class="l">Different winners</td><td>{r["Different winners"]}</td><td>{r["Different winners if random"]}</td></tr>'
            f'<tr><td class="l">Top-3 share of wins</td><td>{r["Top-3 share of wins %"]}%</td><td>{r["Top-3 share if random %"]}%</td></tr></tbody></table>'
            f'<div class="s">{r["Games with a winner"]} games with a winner</div></div>')
    st.markdown('<div class="mnm-metrics" style="grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));">'
                + cards + '</div>', unsafe_allow_html=True)


def show_store_stats(sdf):
    pods = pod_table(sdf)
    reg = pods[pods['wins'] <= 1]
    if len(reg) < 10:
        st.info("Not enough games recorded yet for store stats.")
        return

    # ---- Mondays as a whole ----
    nights = sdf.groupby('Date')['Player'].nunique()
    metric_grid([("Mondays", len(nights)), ("Pods played", len(pods)),
                 ("Avg players / Monday", f"{nights.mean():.1f}"),
                 ("Avg pod size", f"{pods['size'].mean():.2f}")])

    rows = len(sdf)
    W, L, D = (int((sdf['Res'] == r).sum()) for r in ('win', 'loss', 'draw'))
    avg_win = (W + 0.25 * D) / rows * 100
    pts_game = (W * WIN_PTS + L * LOSS_PTS + D * DRAW_PTS) / rows
    metric_grid([("Draw chance (per pod)", f"{(reg['wins'] == 0).mean() * 100:.1f}%"),
                 ("Expected Win% (any game)", f"{avg_win:.1f}%"),
                 ("Avg points / game", f"{pts_game:+.2f}")])
    st.caption(
        "**Draw chance:** how often a game ends with no winner. "
        "**Expected Win%:** the win rate a perfectly average player ends up with. It's above 25% because some pods "
        "only have 3 players, which makes winning easier. "
        f"**Avg points / game:** what the average player earns per game under {WIN_PTS:+d}/{LOSS_PTS:+d}. "
        "Above 0 means the average player slowly gains points just by showing up.")

    t = reg.groupby('size').agg(Pods=('wins', 'size'), Draws=('wins', lambda x: int((x == 0).sum()))).reset_index()
    t['Share'] = (t['Pods'] / t['Pods'].sum() * 100).round(1)
    t['Draw chance'] = (t['Draws'] / t['Pods'] * 100).round(1)
    t['Expected win%'] = ((100 - t['Draw chance']) / t['size']).round(1)
    st.markdown("#### 🎲 By Pod Size")
    html_table(t[['size', 'Pods', 'Share', 'Draw chance', 'Expected win%']],
               headers={'size': 'Pod size', 'Share': 'Share of pods', 'Draw chance': 'Draw chance',
                        'Expected win%': 'Fair win %'},
               fmt={'size': _int, 'Pods': _int, 'Share': _pct, 'Draw chance': _pct, 'Expected win%': _pct},
               left=['size'])
    st.caption("Fair win % is the chance for one player at a table of that size if everyone were equally "
               "good, after allowing for draws.")

    # ---- Seats ----
    st.markdown("#### 💺 Win% by Seat")
    for n in (3, 4):
        st.markdown(f"**{n}-player pods**")
        res = seat_analysis(sdf, n)
        if res is None:
            st.caption("Not enough pods yet.")
            continue
        table, p, total = res
        html_table(table, fmt={'Seat': _int, 'Games': _int, 'Wins': _int, 'Win%': _pct, 'Fair%': _pct,
                               'Edge': _signed}, left=['Seat'])
        st.caption(f"{total} pods. " + (
            f"Seat order looks like it matters: a gap this big would only happen by luck {_odds(p)}." if p < 0.05 else
            "No clear seat advantage: gaps this size are normal luck."))
    st.caption("**Fair%** is what each seat would win if seat order didn't matter. **Edge** is the real win% minus "
               "fair%. Small swings in either direction are normal, especially with few games.")

    # ---- Same winners? ----
    st.markdown("#### 🏆 Do the Same People Keep Winning?")
    wc = winner_concentration(sdf)
    if wc is None:
        st.caption("Not enough decided pods yet.")
    else:
        metric_grid([("Top 5 winners' share of wins", f"{wc['top5'] * 100:.0f}%"),
                     ("Luck-only league", f"{wc['luck_top5'] * 100:.0f}%",
                      f"normal range {wc['luck_lo'] * 100:.0f}-{wc['luck_hi'] * 100:.0f}%"),
                     ("Their share of games played", f"{wc['games_share'] * 100:.0f}%")])
        if wc['p_top5'] < 0.01 or wc['p_chi'] < 0.01:
            verdict = "Very unlikely to be luck: wins are bunched up on certain players"
        elif wc['p_top5'] < 0.05 or wc['p_chi'] < 0.05:
            verdict = "Unlikely to be luck: wins are more bunched up than chance usually gives"
        else:
            verdict = "Looks like normal luck"
        st.markdown(f"**{verdict}.** If every winner were picked at random, the top 5 would usually take about "
                    f"{wc['luck_top5'] * 100:.0f}% of wins (anywhere from {wc['luck_lo'] * 100:.0f}% to "
                    f"{wc['luck_hi'] * 100:.0f}% is normal). A top-5 share this high would happen {_odds(wc['p_top5'])}.")
        with st.expander("How to read this section (no math needed)"):
            st.markdown(
                "**The dice test.** We replay every game from your data with the same players at the same tables, "
                "but pick each winner with a dice roll. We do that 2,000 times to see what a purely lucky league "
                "looks like, then compare it to what really happened.\n\n"
                "- **Top 5 winners' share of wins:** the slice of all wins taken by your five biggest winners.\n"
                "- **Luck-only league:** the slice the dice would give those five. If the real number is well above it, "
                "wins are bunching up.\n"
                "- **Their share of games played:** a fairness check. People who play more should win more, so if the "
                "top 5 also play a lot of the games, part of their share is just attendance. The dice already "
                "account for that.\n"
                "- **The bar chart:** each bar is one player's wins minus the wins dice would give them. Right of zero "
                "means winning more than luck predicts. Green means far enough above that it's unlikely to be a fluke.")

        tb = wc['table'][wc['table']['Pods'] >= MIN_PODS_BREAKDOWN].copy()
        if not tb.empty:
            tb['Verdict'] = np.select([tb['z'] >= 2, tb['z'] <= -2],
                                      ['Beyond luck (above)', 'Beyond luck (below)'], 'Within luck')
            fig = px.bar(tb, x='Above luck', y='Player', orientation='h', color='Verdict', text='Above luck',
                         template="plotly_dark",
                         color_discrete_map={'Beyond luck (above)': '#00ffcc', 'Beyond luck (below)': '#ff6b6b',
                                             'Within luck': '#6b7280'},
                         hover_data={'Wins': True, 'Expected': ':.1f', 'z': ':.2f', 'Verdict': False},
                         labels={'Above luck': 'Wins above what luck predicts'})
            fig.update_traces(texttemplate='%{x:+.1f}', textposition='outside', cliponaxis=False)
            fig.update_layout(yaxis=dict(categoryorder='total ascending', title=None), showlegend=False)
            fig.add_vline(x=0, line_color="gray", opacity=0.5)
            phone_fig(fig, max(260, 28 * len(tb)))
            st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
            st.markdown(
                '<small><span style="color:#00ffcc">Green</span> = far above luck, '
                '<span style="color:#ff6b6b">red</span> = far below, '
                '<span style="color:#9ca3af">gray</span> = within normal luck. '
                f"Each bar is a player's wins minus what luck predicts for the pods they sat in. With {len(tb)} players, "
                f"about {len(tb) * 0.0455:.1f} would land far from zero by chance alone, so treat one bar as a hint "
                "and a cluster of them as evidence.</small>", unsafe_allow_html=True)

        pt = period_table(sdf)
        if not pt.empty:
            st.markdown("**By quarter: are a few people taking most of the wins?**")
            quarter_cards(pt)
            latest = pt.iloc[0]
            st.markdown(
                "Each quarter, we also replay the games with dice: same players, same tables, but a random winner each "
                "game. **Real** is what happened. **Dice** is what dice typically give.\n\n"
                "- **Different winners:** how many people won at least one game. Fewer than the dice number means wins "
                "are landing on fewer people.\n"
                "- **Top-3 share of wins:** the slice of that quarter's wins taken by its three biggest winners. "
                "Higher than the dice number means a few people are dominating.\n"
                "- **The colored label:** our plain-English call. *Same few people winning* means dice would produce a "
                "quarter this lopsided less than 1 time in 20. *Leaning that way* means about 1 time in 5 or less. "
                "*Normal luck* means it looks like an ordinary random quarter.\n\n"
                f"**Example:** in {latest['Period']}, {latest['Different winners']} different players won a game "
                f"(dice would give about {latest['Different winners if random']}), and the top 3 took "
                f"{latest['Top-3 share of wins %']}% of the wins (dice would give about "
                f"{latest['Top-3 share if random %']}%).\n\n"
                "One quarter is only around 50 to 60 games, so a single card can't prove much. The real signal is the "
                "same label showing up quarter after quarter.")

    # ---- Attendance ----
    st.markdown("#### 📅 Monday Attendance")
    att = nights.reset_index()
    att.columns = ['Date', 'Players']
    fig = px.bar(att, x='Date', y='Players', template="plotly_dark")
    fig.update_traces(marker_color='#00ffcc', hovertemplate="%{x|%b %d, %Y}<br>%{y} players<extra></extra>")
    fig.update_layout(xaxis=dict(title=None, tickformat="%b %d<br>%Y", nticks=5), yaxis=dict(title=None))
    phone_fig(fig, 260)
    st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)
    st.caption("Each bar is one Monday: how many different players showed up. Tap a bar for the exact date.")
    big = nights.idxmax()
    drawn = pods[pods['wins'] == 0].groupby('Date').size()
    facts = f"🎉 Biggest night: **{big.strftime('%b %d, %Y')}** with **{nights.max()}** players."
    if not drawn.empty:
        facts += f"  🤝 Most draws in a night: **{drawn.max()}** ({drawn.idxmax().strftime('%b %d, %Y')})."
    st.markdown(facts)


# --- APP START ---
st.markdown(MOBILE_CSS, unsafe_allow_html=True)
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
    except ValueError:
        def_year_idx = 0
    sel_year = st.sidebar.selectbox("Year", available_years, index=def_year_idx)

    scoring_text = f"Win {WIN_PTS:+d}, Loss {LOSS_PTS:+d}, Draw {DRAW_PTS}"
    is_q4 = False

    if sel_year >= CHAMPIONSHIP_START_YEAR:
        q_options = ["Q1", "Q2", "Q3", "Q4 (Championship)"]
        def_q_idx = next((i for i, opt in enumerate(q_options) if curr_q_label in opt), 0)
        sel_q = st.sidebar.selectbox("Quarter", q_options, index=def_q_idx)

        if "Q4" in sel_q:
            is_q4 = True
            year_df = raw_df[raw_df['Year'] == sel_year]
            min_req = MIN_GAMES_YEAR
            rate = Fraction(CARRY_WIN_PTS, WIN_PTS)
            desc = (
                f"🏆 **Championship Mode (Q4)**: Same scoring as Q1-Q3 ({scoring_text}), plus a "
                f"**carry-in** from Q1-Q3 at {rate} rate ({CARRY_WIN_PTS:+d} per win, {CARRY_LOSS_PTS:+d} per loss"
                + (", never below 0" if CARRY_FLOOR_AT_ZERO else "")
                + f"). {min_req} game min, counting all games this year."
            )
        else:
            season_df = raw_df[(raw_df['Year'] == sel_year) & (raw_df['Quarter'] == sel_q[:2])]
            min_req = MIN_GAMES_Q
            desc = (f"📈 **Points Mode ({sel_q[:2]})**: Base {BASE_RATING} Rating. "
                    f"{scoring_text} ({min_req} game min)")
    else:
        season_df = raw_df[raw_df['Year'] == sel_year]
        min_req = MIN_GAMES_LEGACY
        desc = f"📜 **Legacy Mode**: Historic scoring ({scoring_text}). ({min_req} game min)"

    # --- RAGAVAN CHECK ---
    # In Q4 the board still has data (the carry-in) before the first Q4 game is played.
    has_data = (not year_df.empty) if is_q4 else (not season_df.empty)
    if not has_data:
        st.markdown("<h2 style='text-align: center;'>⚔️ The MNM Stats</h2>", unsafe_allow_html=True)
        st.warning(f"🐒 **The data you are looking for was stolen by Ragavan!**")
        st.info(f"There are no recorded matches for **{sel_year}** yet.")
        st.stop()

    if is_q4:
        leaderboard, bench = calculate_q4_stats(year_df, min_req)
    else:
        leaderboard, bench = calculate_league_stats(season_df, min_req)

    # Board as it stood before the latest week, for the rank arrows
    move_src = year_df[year_df['Quarter'] == 'Q4'] if is_q4 else season_df
    moves, move_note = {}, ""
    if not move_src.empty and not leaderboard.empty:
        latest_week = move_src['Year_Week'].max()
        prev_src = (year_df if is_q4 else season_df)
        prev_src = prev_src[prev_src['Year_Week'] < latest_week]
        if not prev_src.empty:
            prev_lb, _ = (calculate_q4_stats(prev_src, min_req) if is_q4
                          else calculate_league_stats(prev_src, min_req))
            moves = week_moves(leaderboard, prev_lb)
            last_day = move_src.loc[move_src['Year_Week'] == latest_week, 'Date'].max()
            move_note = f"Arrows show how each rank moved after the {last_day.strftime('%b %d')} games."

    st.markdown("<h2 style='text-align: center;'>⚔️ The MNM Stats</h2>", unsafe_allow_html=True)
    st.info(desc)
    season_label = f"{sel_year} · {sel_q}" if sel_year >= CHAMPIONSHIP_START_YEAR else f"{sel_year}"
    st.caption(f"Showing **{season_label}**. To change the season, tap the arrow at the top-left of the screen.")

    # Badge data shared by My Stats (earned only) and the Badges tab (everything)
    curr_q_num = int(curr_q_label[1])
    all_earned, badge_extras = compute_badges(raw_df, curr_year, curr_q_num)
    holders = {b[0]: sorted(((p, e[b[0]]) for p, e in all_earned.items() if b[0] in e),
                            key=lambda x: (-x[1], x[0])) for b in BADGES}
    rarity = {k: len(v) / len(all_earned) * 100 for k, v in holders.items()}
    player_names = sorted(raw_df['Player'].unique())

    # Define the top-level tabs
    tab1, tab2, tab_store, tab_badges, tab3 = st.tabs(["🏆 Board", "👤 My Stats", "🏪 Store", "🏅 Badges", "📜 Log"])

    # --- TAB 1: LEADERBOARD ---
    with tab1:
        def row_bg(row):
            if row['Rank'] <= TOP_N:
                return 'rgba(212, 175, 55, 0.32)'
            if row['Rating'] > BASE_RATING:   # silver = positive points overall
                return 'rgba(192, 192, 192, 0.18)'
            return ''

        if not leaderboard.empty:
            start_txt = "your starting score (100 + carry-in)" if is_q4 else f"the {BASE_RATING}-point starting score"
            st.markdown(
                '<div class="mnm-legend">'
                '<span class="mnm-chip" style="background: rgba(212, 175, 55, 0.32);">Gold</span>'
                f'top {TOP_N}<br>'
                '<span class="mnm-chip" style="background: rgba(192, 192, 192, 0.18);">Silver</span>'
                f'Over {BASE_RATING} points</div>',
                unsafe_allow_html=True)
            show_lb = leaderboard.copy()
            if moves:
                show_lb.insert(1, 'Move', [move_badge(*moves[p]) for p in show_lb['Player']])
            html_table(show_lb,
                       headers={'Rank': '#', 'Move': '', 'Carry': 'Carry-in', 'Record': 'W/L/D', 'GP': 'Games'},
                       fmt={'Rank': _int, 'Rating': _int, 'Carry': _int, 'Win%': _pct, 'GP': _int},
                       row_bg=row_bg, left=['Player'], html_cols=('Move',))
            if move_note:
                st.caption(move_note + " NEW = wasn't ranked before.")
        elif is_q4:
            st.caption(f"No one has reached the {int(min_req)}-game minimum yet.")

        # --- THE BENCH ---
        if not bench.empty:
            st.divider()
            st.markdown("### ⏳ The Bench")
            scope = "games this year" if is_q4 else "games"
            st.caption(f"Unranked players (Under {int(min_req)} {scope})")
            html_table(bench,
                       headers={'Would_Be': 'Rank?', 'Carry': 'Carry-in', 'Record': 'W/L/D', 'GP': 'Games'},
                       fmt={'Would_Be': _int, 'Carry': _int, 'GP': _int},
                       left=['Player'])

    # --- TAB 2: MY STATS ---
    with tab2:
        player_sel = st.selectbox("Search Player", player_names)
        p_data = raw_df[raw_df['Player'] == player_sel].copy()
        p_data = p_data.sort_values(['Date', 'Round'])

        # 1. Lifetime Logic
        t_wins = (p_data['Res'] == 'win').sum()
        t_draws = (p_data['Res'] == 'draw').sum()
        total_gp = len(p_data)
        lifetime_wr = ((t_wins + (t_draws * 0.25)) / total_gp * 100) if total_gp > 0 else 0

        # 2. Form Logic (Last 32 games)
        p_data['Win_Val'] = (p_data['Res'] == 'win').astype(int) + ((p_data['Res'] == 'draw').astype(int) * 0.25)
        recent_32 = p_data.tail(32)
        form_32_wr = (recent_32['Win_Val'].mean() * 100) if not recent_32.empty else 0

        # Metrics Display
        metric_grid([("Games Played", total_gp), ("Total WR", f"{lifetime_wr:.1f}%"),
                     ("Last 32 games WR", f"{form_32_wr:.1f}%")])

        # Percentile among players with a similar number of lifetime games
        bracket_label, pct, bracket_n = bracket_percentile(raw_df, player_sel)
        if pct is None:
            st.caption(f"📊 You're the only player with {bracket_label} games, so there's no comparison group yet.")
        else:
            st.markdown(f"📊 Your win rate beats **{pct:.0f}%** of the other {bracket_n - 1} players with **{bracket_label}** games.")

        # Lucky Seat (4-player only)
        lucky_data = p_data[p_data['Pod Size'] == 4]
        best_seat = "N/A"
        if not lucky_data.empty:
            seat_stats = lucky_data.groupby('Seat')['Res'].apply(lambda x: (x == 'win').sum() / len(x))
            if not seat_stats.empty:
                best_seat = f"Seat {int(seat_stats.idxmax())}"

        st.divider()
        col_left, col_right = st.columns(2)
        with col_left:
            st.markdown("#### 💀 Nemesis List")
            my_pod_ids = p_data['Pod ID'].unique()
            pod_winners = raw_df[(raw_df['Pod ID'].isin(my_pod_ids)) & (raw_df['Res'] == 'win') & (raw_df['Player'] != player_sel)]
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
            current_window = wins_list[start: i + 1]
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
            xaxis=dict(showgrid=False, tickformat="%b %d<br>%Y", nticks=5)
        )
        fig.add_hline(y=25, line_dash="dash", line_color="gray", opacity=0.4)
        phone_fig(fig, 300)
        st.plotly_chart(fig, use_container_width=True, config=CHART_CONFIG)

        # --- BADGES (earned only; the full set lives in the Badges tab) ---
        st.divider()
        my_badges = all_earned.get(player_sel, {})
        st.markdown(f"#### 🏅 Badges ({len(my_badges)}/{len(BADGES)})")
        if my_badges:
            earned_cards = "".join(
                f'<div class="mnm-metric" style="border-left:5px solid #00ffcc;">'
                f'<div style="font-size:22px">{icon}</div>'
                f'<b>{html.escape(name)}{f" ×{my_badges[key]}" if my_badges[key] > 1 else ""}</b><br>'
                f'<small style="color:gray">{html.escape(desc)}</small></div>'
                for key, icon, name, desc in sorted((b for b in BADGES if b[0] in my_badges),
                                                    key=lambda b: rarity[b[0]]))
            st.markdown('<div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(155px, 1fr)); '
                        f'grid-auto-rows:1fr; gap:8px;">{earned_cards}</div>', unsafe_allow_html=True)
        else:
            st.caption("No badges yet. Play some games and they'll show up here.")
        # --- NEXT UP: the badges this player is closest to ---
        st.markdown("#### 🎯 Next Up")
        by_key = {b[0]: b for b in BADGES}
        todo = next_up(raw_df, player_sel, my_badges, badge_extras, curr_year, curr_q_num)

        def up_row(it):
            _, icon, name, desc = by_key[it['key']]
            pct = int(round(it['frac'] * 100))
            tag = (f'<span style="font-size:11px; color:#00ffcc; margin-left:6px;">{html.escape(it["tag"])}</span>'
                   if it['tag'] else '')
            return (f'<div style="background: rgba(255,255,255,0.05); padding: 12px 14px; border-radius: 14px; '
                    f'margin-bottom: 8px; border-left: 5px solid #00ffcc;">'
                    f'<div>{icon} <b>{html.escape(name)}</b>{tag}</div>'
                    f'<div style="font-size:13px; color:gray; margin-top:2px;">{html.escape(desc)}</div>'
                    f'<div style="height:8px; border-radius:4px; background: rgba(128,128,128,0.25); margin:10px 0 6px;">'
                    f'<div style="height:8px; border-radius:4px; width:{pct}%; background:#00ffcc;"></div></div>'
                    f'<small style="color:gray">{html.escape(it["note"])}</small></div>')

        if todo:
            st.markdown("".join(up_row(i) for i in todo[:NEXT_UP_SHOW]), unsafe_allow_html=True)
            if len(todo) > NEXT_UP_SHOW:
                st.markdown(f'<details><summary style="cursor:pointer; margin-bottom:8px;">'
                            f'Show {len(todo) - NEXT_UP_SHOW} more</summary>'
                            + "".join(up_row(i) for i in todo[NEXT_UP_SHOW:]) + '</details>',
                            unsafe_allow_html=True)
        else:
            st.caption("Nothing in progress yet. Play a few games and your closest badges will show up here.")
        st.caption("The full list of badges, and who holds each one, is in the 🏅 Badges tab.")

    # --- TAB: BADGES ---
    with tab_badges:
        claimed = [b for b in BADGES if holders[b[0]]]
        st.markdown(f"#### 🏅 {len(claimed)} of {len(BADGES)} badges claimed")
        st.markdown(
            '<div class="mnm-legend">'
            '<span class="mnm-chip" style="background: rgba(0, 255, 204, 0.22);">Highlighted</span>'
            'claimed by at least one player<br>'
            '<span class="mnm-chip" style="background: rgba(128, 128, 128, 0.22);">Greyed</span>'
            'nobody has it yet</div>', unsafe_allow_html=True)
        st.caption("Tap a badge to see who has it.")

        def badge_card(b):
            key, icon, name, desc = b
            who_list = holders[key]
            have = bool(who_list)
            pct_have = rarity[key]
            if not have:
                rar = "Unclaimed"
            else:
                n = len(who_list)
                share = f"{pct_have:.0f}%" if pct_have >= 10 else f"{pct_have:.1f}%"
                rar = f"{n} {'player' if n == 1 else 'players'} · {share}"
            dim = "" if have else "opacity:0.35; filter:grayscale(1);"
            edge = "#00ffcc" if have else "#555"
            bg = "rgba(0, 255, 204, 0.08)" if have else "rgba(255,255,255,0.05)"
            who = ", ".join(html.escape(p) + (f" ×{n}" if n > 1 else "") for p, n in who_list) or "Nobody yet"
            return (f'<details style="box-sizing:border-box; background: {bg}; padding: 22px 20px; border-radius: 14px; border-left: 5px solid {edge};">'
                    f'<summary style="display:block; cursor:pointer;">'
                    f'<div style="min-height:150px;">'
                    f'<div style="{dim}"><div style="font-size:22px">{icon}</div><b>{name}</b><br>'
                    f'<small style="color:gray">{desc}</small>'
                    f'<div style="margin-top:6px;"><small>🎯 {rar}</small></div></div>'
                    f'</div></summary>'
                    f'<div style="margin-top:8px; font-size:13px;"><small style="color:gray">Earned by:</small><br>{who}</div>'
                    f'</details>')

        # claimed first (rarest on top), then unclaimed in list order
        ordered_badges = sorted(BADGES, key=lambda b: (not holders[b[0]], rarity[b[0]] if holders[b[0]] else 0))
        st.markdown('<div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(155px, 1fr)); '
                    'gap:16px; align-items:stretch;">'
                    + "".join(badge_card(b) for b in ordered_badges) + '</div>',
                    unsafe_allow_html=True)

    # --- TAB: STORE STATS ---
    with tab_store:
        scope_options = ["All time"] + available_years
        scope = st.selectbox("Season", scope_options, key="store_scope",
                             index=scope_options.index(curr_year) if curr_year in scope_options else 0)
        store_df = raw_df if scope == "All time" else raw_df[raw_df['Year'] == scope]
        show_store_stats(store_df[['Date', 'Round', 'Pod ID', 'Player', 'Res', 'Year', 'Quarter', 'Seat']])

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
                    winners = pod[pod['Res'] == 'win']['Player'].tolist()
                    st.markdown(f"""<div style="background: rgba(255,255,255,0.05); padding: 10px; border-radius: 10px; margin-bottom: 5px; border-left: 5px solid #00ffcc;">
                        <b>{pod['Date'].iloc[0].strftime('%m/%d')} Winner: {", ".join(winners) if winners else "Draw"}</b><br>
                        <small>{", ".join(pod['Player'].tolist())}</small></div>""", unsafe_allow_html=True)