import discord
from discord import ui, Embed, Color, app_commands
from discord.ui import Button, View, Select, Modal, TextInput
import os
import json
import re
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from dotenv import load_dotenv
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import asyncio

# Toronto timezone
TORONTO_TZ = ZoneInfo("America/Toronto")

load_dotenv()
DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")

DATA_FILE = "bankroll_data.json"

intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)
tree = app_commands.CommandTree(client)

# Store active bet sessions
active_sessions = {}

# ---------- PRIVATE MODALS FOR BET ENTRY ----------

class BetDataModal(Modal, title="📊 BET DETAILS"):
    """Private modal for entering bet description, odds, and stake"""
    description_input = TextInput(label="Pick/Description", placeholder="e.g., Lakers -5.5 or Over 220", required=True, max_length=100)
    odds_input = TextInput(label="Odds (Decimal)", placeholder="e.g., 1.85", required=True, max_length=10)
    stake_input = TextInput(label="Stake ($)", placeholder="e.g., 50", required=True, max_length=20)
    
    def __init__(self, user_id: str, sport: str, bet_type: str, pending: bool = False):
        super().__init__()
        self.user_id = user_id
        self.sport = sport
        self.bet_type = bet_type
        self.pending = pending
    
    async def on_submit(self, interaction: discord.Interaction):
        try:
            description = self.description_input.value.strip()
            odds = float(self.odds_input.value.replace(",", "."))
            stake = float(self.stake_input.value.replace("$", "").replace(",", "."))
            
            # Validation
            if odds <= 1:
                embed = Embed(title="❌ Invalid Odds", description="Odds must be greater than 1.", color=COLORS["red"])
                await interaction.response.send_message(embed=embed, ephemeral=True)
                return
            
            if stake <= 0:
                embed = Embed(title="❌ Invalid Stake", description="Stake must be positive.", color=COLORS["red"])
                await interaction.response.send_message(embed=embed, ephemeral=True)
                return
            
            data = get_user_data(self.user_id)
            user = data[self.user_id]
            
            if self.pending:
                # Save as pending bet
                pending_bet = {
                    "date": datetime.now(TORONTO_TZ).strftime("%Y-%m-%d"),
                    "sport": self.sport,
                    "bet_type": self.bet_type,
                    "description": description,
                    "odds": odds,
                    "stake": stake
                }
                user["pending_bets"].append(pending_bet)
                update_user_data(self.user_id, user)
                
                potential = stake * (odds - 1)
                sport_emoji = SPORT_EMOJIS.get(self.sport, "🎯")
                
                embed = Embed(title="⏳ PENDING BET ADDED", color=COLORS["orange"])
                embed.add_field(
                    name="📋 Details",
                    value=f"{sport_emoji} **{self.sport}** | {self.bet_type}\n"
                          f"📝 {description}\n"
                          f"📊 Odds: **{odds}**\n"
                          f"💵 Stake: **${stake:.2f}**",
                    inline=False
                )
                embed.add_field(name="💰 Potential Win", value=f"**+${potential:.2f}**", inline=False)
                embed.set_footer(text="Use /menu → Pending Bets → Settle when it's done!")
                await interaction.response.send_message(embed=embed, ephemeral=True)
            else:
                # Store in session and ask for result
                data = get_user_data(self.user_id)
                user = data[self.user_id]
                
                active_sessions[self.user_id]["description"] = description
                active_sessions[self.user_id]["odds"] = odds
                active_sessions[self.user_id]["stake"] = stake
                
                await interaction.response.send_message(
                    embed=Embed(title="🏆 RESULT", description="Select the outcome:", color=COLORS["blue"]),
                    view=ResultSelectView(self.user_id),
                    ephemeral=True
                )
        
        except ValueError:
            embed = Embed(title="❌ Invalid Input", description="Check your odds and stake values!", color=COLORS["red"])
            await interaction.response.send_message(embed=embed, ephemeral=True)


class GoalModal(Modal, title="🎯 PROFIT GOAL"):
    """Private modal for setting profit goals"""
    goal_input = TextInput(label="Target Profit ($)", placeholder="e.g., 500 for $500", required=True, max_length=20)
    
    def __init__(self, user_id: str):
        super().__init__()
        self.user_id = user_id
    
    async def on_submit(self, interaction: discord.Interaction):
        try:
            goal = float(self.goal_input.value.replace("$", "").replace(",", ""))
            
            if goal <= 0:
                embed = Embed(title="❌ Invalid", description="Goal must be positive.", color=COLORS["red"])
                await interaction.response.send_message(embed=embed, ephemeral=True)
                return
            
            data = get_user_data(self.user_id)
            user = data[self.user_id]
            user["profit_goal"] = goal
            update_user_data(self.user_id, user)
            
            progress = get_goal_progress(user)
            bar = create_progress_bar(progress["progress"])
            
            embed = Embed(title="🎯 GOAL SET!", color=COLORS["gold"])
            embed.add_field(name="Target", value=f"**${goal:.2f}** profit", inline=True)
            embed.add_field(name="Current", value=f"**${user['total_profit']:.2f}**", inline=True)
            embed.add_field(name="Progress", value=f"{bar} **{progress['progress']:.1f}%**", inline=False)
            await interaction.response.send_message(embed=embed, ephemeral=True)
        
        except ValueError:
            embed = Embed(title="❌ Invalid", description="Enter a number like `500`", color=COLORS["red"])
            await interaction.response.send_message(embed=embed, ephemeral=True)

# Allowed channel for bot usage (case-insensitive, spaces/dashes flexible)
ALLOWED_CHANNEL = "bankroll bot"

BET_TYPES = ["Straight", "Parlay", "2-Man Parlay", "3-Man Parlay", "Round Robin", "Teaser"]
SPORTS = ["NBA", "NHL", "NFL", "MLB", "Soccer", "UFC/MMA", "Tennis", "Other"]
SPORT_EMOJIS = {"NBA": "🏀", "NHL": "🏒", "NFL": "🏈", "MLB": "⚾", "Soccer": "⚽", "UFC/MMA": "🥊", "Tennis": "🎾", "Other": "🎯"}

def is_allowed_channel(channel_name: str) -> bool:
    """Check if the current channel is the allowed bankroll channel"""
    # Just check if channel contains 'bankroll' (case-insensitive)
    # This handles both "bankroll-bot" and "🎰-bankroll-bot" formats
    return "bankroll" in channel_name.lower()

# Colors for embeds
COLORS = {
    "green": Color.green(),
    "red": Color.red(),
    "blue": Color.blue(),
    "gold": Color.gold(),
    "purple": Color.purple(),
    "orange": Color.orange(),
    "gray": Color.greyple()
}


# ---------- DATA FUNCTIONS ----------

def load_data():
    if not os.path.exists(DATA_FILE):
        return {}
    try:
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)


def get_user_data(user_id: str):
    data = load_data()
    if user_id not in data:
        data[user_id] = {
            "starting_bankroll": 0.0,
            "current_bankroll": 0.0,
            "total_staked": 0.0,
            "total_profit": 0.0,
            "wins": 0,
            "losses": 0,
            "pushes": 0,
            "bets": [],
            "pending_bets": [],
            "daily_limit": None,
            "bets_today": 0,
            "last_date": str(date.today()),
            "profit_goal": None,
            "loss_alert_percent": 20,
            "daily_starting_bankroll": 0.0,
            "last_summary_date": None
        }
        save_data(data)
    else:
        # Ensure new fields exist for existing users
        user = data[user_id]
        if "pending_bets" not in user:
            user["pending_bets"] = []
        if "profit_goal" not in user:
            user["profit_goal"] = None
        if "loss_alert_percent" not in user:
            user["loss_alert_percent"] = 20
        if "daily_starting_bankroll" not in user:
            user["daily_starting_bankroll"] = user["current_bankroll"]
        if "last_summary_date" not in user:
            user["last_summary_date"] = None
        if "win_streak" not in user:
            user["win_streak"] = 0
        if "loss_streak" not in user:
            user["loss_streak"] = 0
        if "best_win" not in user:
            user["best_win"] = 0.0
        if "worst_loss" not in user:
            user["worst_loss"] = 0.0
        if "username" not in user:
            user["username"] = "Unknown"
        save_data(data)
    return data


def update_user_data(user_id: str, user_data: dict):
    data = load_data()
    data[user_id] = user_data
    save_data(data)


def reset_if_new_day(user_data: dict):
    today_str = str(date.today())
    if user_data.get("last_date") != today_str:
        user_data["last_date"] = today_str
        user_data["bets_today"] = 0
        user_data["daily_starting_bankroll"] = user_data["current_bankroll"]
    return user_data


def check_loss_alert(user_data: dict):
    """Check if user has lost too much today"""
    daily_start = user_data.get("daily_starting_bankroll", user_data["current_bankroll"])
    if daily_start <= 0:
        return None
    
    current = user_data["current_bankroll"]
    loss_percent = ((daily_start - current) / daily_start) * 100
    alert_threshold = user_data.get("loss_alert_percent", 20)
    
    if loss_percent >= alert_threshold:
        return loss_percent
    return None


def get_goal_progress(user_data: dict):
    """Get profit goal progress"""
    goal = user_data.get("profit_goal")
    if not goal:
        return None
    
    profit = user_data["total_profit"]
    progress = (profit / goal) * 100 if goal > 0 else 0
    return {"goal": goal, "current": profit, "progress": min(progress, 100)}


def get_achievements(user_data: dict) -> list:
    """Calculate user achievements/badges"""
    achievements = []
    profit = user_data["total_profit"]
    wins = user_data["wins"]
    losses = user_data["losses"]
    total = wins + losses
    
    # Milestone achievements
    if profit >= 100:
        achievements.append(("🟢", "Century Club", f"+${profit:.0f} profit"))
    if profit >= 500:
        achievements.append(("🔵", "High Roller", f"+${profit:.0f} profit"))
    if profit >= 1000:
        achievements.append(("⭐", "Legend", f"+${profit:.0f} profit"))
    
    # Win rate achievements
    if total >= 10:
        wr = (wins / total * 100)
        if wr >= 60:
            achievements.append(("🎯", "Sharp Bettor", f"{wr:.1f}% win rate"))
        if wr >= 70:
            achievements.append(("💎", "Elite Bettor", f"{wr:.1f}% win rate"))
    
    # Volume achievements
    if wins >= 10:
        achievements.append(("🏆", "Victory Counter", f"{wins} wins"))
    if wins >= 25:
        achievements.append(("👑", "Win Master", f"{wins} wins"))
    
    # Streak achievements
    if user_data.get("win_streak", 0) >= 3:
        achievements.append(("🔥", "Hot Streak", f"{user_data['win_streak']} W streak"))
    if user_data.get("win_streak", 0) >= 5:
        achievements.append(("🌟", "Blazing", f"{user_data['win_streak']} W streak"))
    
    # First bet achievement
    if total >= 1:
        achievements.append(("🎰", "Rookie", "First bet placed"))
    
    return achievements[:8]  # Max 8 achievements shown


def get_leaderboard(guild_id: str = None) -> list:
    """Get top 5 users by profit"""
    data = load_data()
    users = []
    
    for user_id, user_data in data.items():
        if user_data.get("starting_bankroll", 0) > 0:
            total_bets = user_data["wins"] + user_data["losses"] + user_data.get("pushes", 0)
            wr = (user_data["wins"] / total_bets * 100) if total_bets > 0 else 0
            users.append({
                "id": user_id,
                "name": user_data.get("username", "Unknown"),
                "profit": user_data["total_profit"],
                "wins": user_data["wins"],
                "record": f"{user_data['wins']}-{user_data['losses']}",
                "wr": wr
            })
    
    users.sort(key=lambda x: x["profit"], reverse=True)
    return users[:5]


# ---------- EMBED BUILDERS ----------

def create_menu_embed(user_id: str):
    data = get_user_data(user_id)
    user = data[user_id]
    
    embed = Embed(
        title="🎰 BANKROLL MANAGER",
        description="Track your bets, manage your bankroll, achieve your goals!",
        color=COLORS["purple"]
    )
    
    embed.add_field(
        name="💰 Bankroll",
        value=f"**${user['current_bankroll']:.2f}**",
        inline=True
    )
    embed.add_field(
        name="📈 Profit",
        value=f"**${user['total_profit']:+.2f}**",
        inline=True
    )
    embed.add_field(
        name="🏆 Record",
        value=f"**{user['wins']}W - {user['losses']}L**",
        inline=True
    )
    
    pending_count = len(user.get("pending_bets", []))
    if pending_count > 0:
        embed.add_field(
            name="⏳ Pending",
            value=f"**{pending_count} bet(s)**",
            inline=True
        )
    
    goal_info = get_goal_progress(user)
    if goal_info:
        bar = create_progress_bar(goal_info["progress"])
        embed.add_field(
            name="🎯 Goal Progress",
            value=f"{bar} **{goal_info['progress']:.1f}%**\n${goal_info['current']:.2f} / ${goal_info['goal']:.2f}",
            inline=False
        )
    
    embed.set_footer(text=f"Toronto Time: {datetime.now(TORONTO_TZ).strftime('%H:%M')}")
    return embed


def create_progress_bar(percent, length=10):
    filled = int(percent / 100 * length)
    empty = length - filled
    return "█" * filled + "░" * empty


def create_summary_embed(user_id: str):
    data = get_user_data(user_id)
    user = data[user_id]
    user = reset_if_new_day(user)
    update_user_data(user_id, user)

    starting = user["starting_bankroll"]
    current = user["current_bankroll"]
    total_staked = user["total_staked"]
    total_profit = user["total_profit"]
    wins = user["wins"]
    losses = user["losses"]
    pushes = user.get("pushes", 0)
    bets_count = wins + losses + pushes

    roi = (total_profit / total_staked * 100) if total_staked > 0 else 0.0
    winrate = (wins / (wins + losses) * 100) if (wins + losses) > 0 else 0.0

    color = COLORS["green"] if total_profit >= 0 else COLORS["red"]
    
    embed = Embed(
        title="📊 BANKROLL SUMMARY",
        color=color
    )
    
    embed.add_field(
        name="💰 Bankroll",
        value=f"Starting: **${starting:.2f}**\nCurrent: **${current:.2f}**\nProfit: **${total_profit:+.2f}**",
        inline=True
    )
    
    embed.add_field(
        name="📈 Performance",
        value=f"Staked: **${total_staked:.2f}**\nROI: **{roi:+.2f}%**\nWin Rate: **{winrate:.1f}%**",
        inline=True
    )
    
    embed.add_field(
        name="🏆 Record",
        value=f"**{wins}W - {losses}L - {pushes}P**\nTotal: **{bets_count} bets**",
        inline=True
    )
    
    limit_txt = f"{user['daily_limit']}/day" if user["daily_limit"] else "None"
    embed.add_field(
        name="📆 Today",
        value=f"Bets: **{user['bets_today']}**\nLimit: **{limit_txt}**",
        inline=True
    )
    
    goal_info = get_goal_progress(user)
    if goal_info:
        bar = create_progress_bar(goal_info["progress"])
        embed.add_field(
            name="🎯 Goal Progress",
            value=f"{bar}\n**{goal_info['progress']:.1f}%** (${goal_info['current']:.2f}/${goal_info['goal']:.2f})",
            inline=True
        )
    
    pending_count = len(user.get("pending_bets", []))
    embed.add_field(
        name="⏳ Pending Bets",
        value=f"**{pending_count}** awaiting result",
        inline=True
    )
    
    embed.set_footer(text=f"Generated: {datetime.now(TORONTO_TZ).strftime('%Y-%m-%d %H:%M')} (Toronto)")
    return embed


def create_bet_recorded_embed(bet_data: dict, user: dict, result: str):
    stake = bet_data["stake"]
    odds = bet_data["odds"]
    profit = bet_data["profit"]
    
    if result == "w":
        color = COLORS["green"]
        title = "🎉 WIN!"
        total_return = stake + profit
        return_text = f"💰 Return: **${total_return:.2f}** (+${profit:.2f})"
    elif result == "l":
        color = COLORS["red"]
        title = "😔 LOSS"
        return_text = f"💸 Lost: **${stake:.2f}**"
    else:
        color = COLORS["gray"]
        title = "➖ PUSH"
        return_text = f"↩️ Refunded: **${stake:.2f}**"
    
    sport_emoji = SPORT_EMOJIS.get(bet_data["sport"], "🎯")
    
    embed = Embed(title=title, color=color)
    
    embed.add_field(
        name="📋 Bet Details",
        value=f"{sport_emoji} **{bet_data['sport']}** | {bet_data['bet_type']}\n"
              f"📝 {bet_data['description']}\n"
              f"📊 Odds: **{odds}**",
        inline=False
    )
    
    embed.add_field(
        name="💵 Stake",
        value=f"**${stake:.2f}**",
        inline=True
    )
    
    embed.add_field(
        name="💰 Result",
        value=return_text,
        inline=True
    )
    
    embed.add_field(
        name="━━━━━━━━━━━━━━━",
        value=f"🏦 Bankroll: **${user['current_bankroll']:.2f}**\n"
              f"📈 Record: **{user['wins']}W - {user['losses']}L**",
        inline=False
    )
    
    return embed


def create_pending_bet_embed(bet: dict, index: int):
    sport_emoji = SPORT_EMOJIS.get(bet["sport"], "🎯")
    
    embed = Embed(
        title=f"⏳ Pending Bet #{index + 1}",
        color=COLORS["orange"]
    )
    
    embed.add_field(
        name="📋 Details",
        value=f"{sport_emoji} **{bet['sport']}** | {bet['bet_type']}\n"
              f"📝 {bet['description']}\n"
              f"📊 Odds: **{bet['odds']}**\n"
              f"💵 Stake: **${bet['stake']:.2f}**",
        inline=False
    )
    
    potential_win = bet['stake'] * (bet['odds'] - 1)
    embed.add_field(
        name="💰 Potential Win",
        value=f"**+${potential_win:.2f}**",
        inline=True
    )
    
    embed.add_field(
        name="📅 Placed",
        value=bet.get("date", "N/A"),
        inline=True
    )
    
    return embed


def create_loss_alert_embed(loss_percent: float, user: dict):
    embed = Embed(
        title="⚠️ LOSS ALERT",
        description=f"You're down **{loss_percent:.1f}%** today!\n\nTake a break and protect your bankroll.",
        color=COLORS["red"]
    )
    
    daily_start = user.get("daily_starting_bankroll", user["current_bankroll"])
    current = user["current_bankroll"]
    daily_loss = daily_start - current
    
    embed.add_field(
        name="📉 Today's Stats",
        value=f"Started: **${daily_start:.2f}**\nCurrent: **${current:.2f}**\nDown: **${daily_loss:.2f}**",
        inline=False
    )
    
    embed.set_footer(text="Consider taking a break. Betting responsibly is key!")
    return embed


def create_daily_summary_embed(user_id: str):
    data = get_user_data(user_id)
    user = data[user_id]
    
    # Get today's bets
    today_str = date.today().isoformat()
    today_bets = [b for b in user["bets"] if b.get("date", "").startswith(today_str)]
    
    wins = sum(1 for b in today_bets if b["result"] == "w")
    losses = sum(1 for b in today_bets if b["result"] == "l")
    pushes = sum(1 for b in today_bets if b["result"] == "p")
    daily_profit = sum(b["profit"] for b in today_bets)
    
    color = COLORS["green"] if daily_profit >= 0 else COLORS["red"]
    
    embed = Embed(
        title="📅 DAILY SUMMARY",
        description=f"**{datetime.now(TORONTO_TZ).strftime('%A, %B %d, %Y')}**",
        color=color
    )
    
    embed.add_field(
        name="📊 Today's Results",
        value=f"Record: **{wins}W - {losses}L - {pushes}P**\n"
              f"Profit: **${daily_profit:+.2f}**\n"
              f"Bets: **{len(today_bets)}**",
        inline=True
    )
    
    embed.add_field(
        name="💰 Bankroll",
        value=f"Current: **${user['current_bankroll']:.2f}**\n"
              f"Total Profit: **${user['total_profit']:+.2f}**",
        inline=True
    )
    
    pending_count = len(user.get("pending_bets", []))
    if pending_count > 0:
        pending_stake = sum(b["stake"] for b in user["pending_bets"])
        embed.add_field(
            name="⏳ Pending",
            value=f"**{pending_count}** bets (${pending_stake:.2f} at risk)",
            inline=True
        )
    
    goal_info = get_goal_progress(user)
    if goal_info:
        bar = create_progress_bar(goal_info["progress"])
        embed.add_field(
            name="🎯 Goal Progress",
            value=f"{bar} **{goal_info['progress']:.1f}%**",
            inline=True
        )
    
    embed.set_footer(text="Keep grinding! 💪")
    return embed


# ---------- EXCEL EXPORT ----------

def export_to_excel(user_id: str, username: str):
    data = get_user_data(user_id)
    user = data[user_id]
    
    wb = openpyxl.Workbook()
    
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    title_font = Font(bold=True, size=16)
    green_font = Font(bold=True, color="228B22")
    red_font = Font(bold=True, color="CC0000")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center = Alignment(horizontal="center", vertical="center")
    
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    
    ws_dash["A1"] = f"BANKROLL TRACKER - {username.upper()}"
    ws_dash["A1"].font = Font(bold=True, size=18)
    ws_dash.merge_cells("A1:D1")
    
    toronto_now = datetime.now(TORONTO_TZ)
    ws_dash["A2"] = f"Report Date: {toronto_now.strftime('%B %d, %Y at %H:%M')} (Toronto)"
    ws_dash["A2"].font = Font(italic=True, size=10, color="666666")
    
    total_bets = user["wins"] + user["losses"] + user.get("pushes", 0)
    roi = (user["total_profit"] / user["total_staked"] * 100) if user["total_staked"] > 0 else 0
    winrate = (user["wins"] / total_bets * 100) if total_bets > 0 else 0
    
    stats = [
        ("BANKROLL", ""),
        ("Starting", f"${user['starting_bankroll']:.2f}"),
        ("Current", f"${user['current_bankroll']:.2f}"),
        ("Profit/Loss", f"${user['total_profit']:+.2f}"),
        ("", ""),
        ("PERFORMANCE", ""),
        ("Total Staked", f"${user['total_staked']:.2f}"),
        ("ROI", f"{roi:+.2f}%"),
        ("Win Rate", f"{winrate:.1f}%"),
        ("", ""),
        ("RECORD", ""),
        ("Wins", user['wins']),
        ("Losses", user['losses']),
        ("Pushes", user.get('pushes', 0)),
        ("Total Bets", total_bets),
        ("", ""),
        ("GOALS", ""),
        ("Profit Goal", f"${user.get('profit_goal', 0):.2f}" if user.get('profit_goal') else "Not set"),
        ("Pending Bets", len(user.get("pending_bets", []))),
    ]
    
    for i, (label, value) in enumerate(stats, start=4):
        cell_a = ws_dash[f"A{i}"]
        cell_b = ws_dash[f"B{i}"]
        
        if value == "":
            cell_a.value = label
            cell_a.font = Font(bold=True, size=12, color="2F5496")
        else:
            cell_a.value = label
            cell_a.border = border
            cell_b.value = value
            cell_b.border = border
            cell_b.alignment = Alignment(horizontal="right")
            
            if label == "Profit/Loss":
                cell_b.font = green_font if user["total_profit"] >= 0 else red_font
    
    ws_dash.column_dimensions["A"].width = 18
    ws_dash.column_dimensions["B"].width = 15
    
    ws_bets = wb.create_sheet("Bet History")
    
    headers = ["#", "Date", "Sport", "Bet Type", "Pick/Description", "Odds", "Stake", "Result", "Profit/Loss"]
    for col, header in enumerate(headers, start=1):
        cell = ws_bets.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    
    for row, bet in enumerate(user["bets"], start=2):
        ws_bets.cell(row=row, column=1, value=row-1).border = border
        ws_bets.cell(row=row, column=1).alignment = center
        
        ws_bets.cell(row=row, column=2, value=bet.get("date", "N/A")).border = border
        ws_bets.cell(row=row, column=3, value=bet.get("sport", "N/A")).border = border
        ws_bets.cell(row=row, column=4, value=bet.get("bet_type", "Straight")).border = border
        ws_bets.cell(row=row, column=5, value=bet.get("description", "")).border = border
        ws_bets.cell(row=row, column=6, value=bet.get("odds", 0)).border = border
        ws_bets.cell(row=row, column=7, value=f"${bet.get('stake', 0):.2f}").border = border
        
        result = bet.get("result", "").upper()
        result_cell = ws_bets.cell(row=row, column=8, value="WIN" if result == "W" else "LOSS" if result == "L" else "PUSH")
        result_cell.border = border
        result_cell.alignment = center
        if result == "W":
            result_cell.font = green_font
            result_cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        elif result == "L":
            result_cell.font = red_font
            result_cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
        profit = bet.get("profit", 0)
        profit_cell = ws_bets.cell(row=row, column=9, value=f"${profit:+.2f}")
        profit_cell.border = border
        profit_cell.alignment = Alignment(horizontal="right")
        profit_cell.font = green_font if profit >= 0 else red_font
    
    widths = [5, 12, 12, 14, 35, 8, 10, 10, 12]
    for i, w in enumerate(widths, start=1):
        ws_bets.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    
    # Pending Bets Sheet
    ws_pending = wb.create_sheet("Pending Bets")
    ws_pending["A1"] = "PENDING BETS"
    ws_pending["A1"].font = title_font
    
    pending_headers = ["#", "Date", "Sport", "Bet Type", "Pick", "Odds", "Stake", "Potential Win"]
    for col, header in enumerate(pending_headers, start=1):
        cell = ws_pending.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
        cell.border = border
        cell.alignment = center
    
    for row, bet in enumerate(user.get("pending_bets", []), start=4):
        ws_pending.cell(row=row, column=1, value=row-3).border = border
        ws_pending.cell(row=row, column=2, value=bet.get("date", "N/A")).border = border
        ws_pending.cell(row=row, column=3, value=bet.get("sport", "N/A")).border = border
        ws_pending.cell(row=row, column=4, value=bet.get("bet_type", "Straight")).border = border
        ws_pending.cell(row=row, column=5, value=bet.get("description", "")).border = border
        ws_pending.cell(row=row, column=6, value=bet.get("odds", 0)).border = border
        ws_pending.cell(row=row, column=7, value=f"${bet.get('stake', 0):.2f}").border = border
        potential = bet.get("stake", 0) * (bet.get("odds", 1) - 1)
        ws_pending.cell(row=row, column=8, value=f"+${potential:.2f}").border = border
    
    for i, w in enumerate([5, 12, 12, 14, 30, 8, 10, 12], start=1):
        ws_pending.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    
    # By Sport Sheet
    ws_sport = wb.create_sheet("By Sport")
    ws_sport["A1"] = "BREAKDOWN BY SPORT"
    ws_sport["A1"].font = title_font
    
    sport_headers = ["Sport", "Bets", "Wins", "Losses", "Win %", "Profit/Loss", "ROI"]
    for col, header in enumerate(sport_headers, start=1):
        cell = ws_sport.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = center
    
    sport_stats = {}
    for bet in user["bets"]:
        sport = bet.get("sport", "Other")
        if sport not in sport_stats:
            sport_stats[sport] = {"bets": 0, "wins": 0, "losses": 0, "profit": 0, "staked": 0}
        sport_stats[sport]["bets"] += 1
        sport_stats[sport]["staked"] += bet.get("stake", 0)
        sport_stats[sport]["profit"] += bet.get("profit", 0)
        if bet.get("result") == "w":
            sport_stats[sport]["wins"] += 1
        elif bet.get("result") == "l":
            sport_stats[sport]["losses"] += 1
    
    row = 4
    for sport, stats in sorted(sport_stats.items()):
        win_pct = (stats["wins"] / stats["bets"] * 100) if stats["bets"] > 0 else 0
        roi = (stats["profit"] / stats["staked"] * 100) if stats["staked"] > 0 else 0
        
        ws_sport.cell(row=row, column=1, value=sport).border = border
        ws_sport.cell(row=row, column=2, value=stats["bets"]).border = border
        ws_sport.cell(row=row, column=3, value=stats["wins"]).border = border
        ws_sport.cell(row=row, column=4, value=stats["losses"]).border = border
        ws_sport.cell(row=row, column=5, value=f"{win_pct:.1f}%").border = border
        
        profit_cell = ws_sport.cell(row=row, column=6, value=f"${stats['profit']:+.2f}")
        profit_cell.border = border
        profit_cell.font = green_font if stats["profit"] >= 0 else red_font
        
        roi_cell = ws_sport.cell(row=row, column=7, value=f"{roi:+.1f}%")
        roi_cell.border = border
        roi_cell.font = green_font if roi >= 0 else red_font
        
        row += 1
    
    for i, w in enumerate([12, 8, 8, 8, 10, 12, 10], start=1):
        ws_sport.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    
    filename = f"bankroll_{username}_{datetime.now(TORONTO_TZ).strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)
    return filename


# ---------- DISCORD UI VIEWS ----------

class MainMenuView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=300)
        self.user_id = user_id
    
    @ui.button(label="New Bet", style=discord.ButtonStyle.green, emoji="🎲", row=0)
    async def new_bet(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = Embed(title="🎲 NEW BET", description="Select a sport:", color=COLORS["blue"])
        await interaction.response.send_message(embed=embed, view=SportSelectView(self.user_id), ephemeral=True)
    
    @ui.button(label="Pending Bets", style=discord.ButtonStyle.secondary, emoji="⏳", row=0)
    async def pending(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        await interaction.response.send_message(embed=Embed(title="⏳ PENDING BETS", description="Choose an action:", color=COLORS["orange"]), view=PendingMenuView(self.user_id), ephemeral=True)
    
    @ui.button(label="Summary", style=discord.ButtonStyle.blurple, emoji="📊", row=0)
    async def summary(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = create_summary_embed(self.user_id)
        await interaction.response.send_message(embed=embed, ephemeral=True)
    
    @ui.button(label="Daily Report", style=discord.ButtonStyle.blurple, emoji="📅", row=0)
    async def daily(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = create_daily_summary_embed(self.user_id)
        await interaction.response.send_message(embed=embed, ephemeral=True)
    
    @ui.button(label="Export Excel", style=discord.ButtonStyle.gray, emoji="📁", row=1)
    async def export(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        await interaction.response.defer(ephemeral=True)
        try:
            filename = export_to_excel(self.user_id, interaction.user.name)
            await interaction.followup.send("✅ **Report ready!**", file=discord.File(filename), ephemeral=True)
            os.remove(filename)
        except Exception as e:
            print(f"Export error: {e}")
            await interaction.followup.send("❌ Export failed.", ephemeral=True)
    
    @ui.button(label="Set Goal", style=discord.ButtonStyle.gray, emoji="🎯", row=1)
    async def set_goal(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        # PRIVATE MODAL for goal (no public chat!)
        modal = GoalModal(self.user_id)
        await interaction.response.send_modal(modal)
    
    @ui.button(label="Loss Alert", style=discord.ButtonStyle.gray, emoji="⚠️", row=1)
    async def loss_alert(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = Embed(title="⚠️ LOSS ALERT SETTINGS", description="Get warned when you lose too much in a day:", color=COLORS["orange"])
        await interaction.response.send_message(embed=embed, view=LossAlertView(self.user_id), ephemeral=True)
    
    @ui.button(label="Set Limit", style=discord.ButtonStyle.gray, emoji="🚧", row=1)
    async def set_limit(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = Embed(title="🚧 DAILY BET LIMIT", description="How many bets per day?", color=COLORS["blue"])
        await interaction.response.send_message(embed=embed, view=LimitSelectView(self.user_id), ephemeral=True)
    
    @ui.button(label="Reset", style=discord.ButtonStyle.red, emoji="🗑️", row=1)
    async def reset(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = Embed(title="⚠️ RESET ALL DATA?", description="This cannot be undone!", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, view=ConfirmResetView(self.user_id), ephemeral=True)


class PendingMenuView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=120)
        self.user_id = user_id
    
    @ui.button(label="Add Pending Bet", style=discord.ButtonStyle.green, emoji="➕")
    async def add_pending(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        active_sessions[self.user_id] = {"pending": True}
        embed = Embed(title="➕ ADD PENDING BET", description="Select a sport:", color=COLORS["orange"])
        await interaction.response.send_message(embed=embed, view=SportSelectView(self.user_id, pending=True), ephemeral=True)
    
    @ui.button(label="View Pending", style=discord.ButtonStyle.blurple, emoji="👀")
    async def view_pending(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        data = get_user_data(self.user_id)
        user = data[self.user_id]
        pending = user.get("pending_bets", [])
        
        if not pending:
            embed = Embed(title="⏳ PENDING BETS", description="No pending bets!", color=COLORS["gray"])
            await interaction.response.send_message(embed=embed, ephemeral=True)
            return
        
        embed = Embed(title=f"⏳ PENDING BETS ({len(pending)})", color=COLORS["orange"])
        total_stake = 0
        total_potential = 0
        
        for i, bet in enumerate(pending):
            sport_emoji = SPORT_EMOJIS.get(bet["sport"], "🎯")
            potential = bet["stake"] * (bet["odds"] - 1)
            total_stake += bet["stake"]
            total_potential += potential
            
            embed.add_field(
                name=f"#{i+1} {sport_emoji} {bet['sport']}",
                value=f"📝 {bet['description']}\n💵 ${bet['stake']:.2f} @ {bet['odds']}\n💰 +${potential:.2f}",
                inline=True
            )
        
        embed.set_footer(text=f"Total at risk: ${total_stake:.2f} | Potential: +${total_potential:.2f}")
        await interaction.response.send_message(embed=embed, ephemeral=True)
    
    @ui.button(label="Settle Bet", style=discord.ButtonStyle.secondary, emoji="✅")
    async def settle_pending(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        data = get_user_data(self.user_id)
        user = data[self.user_id]
        pending = user.get("pending_bets", [])
        
        if not pending:
            embed = Embed(title="❌ No Pending Bets", description="Nothing to settle!", color=COLORS["gray"])
            await interaction.response.send_message(embed=embed, ephemeral=True)
            return
        
        embed = Embed(title="✅ SETTLE BET", description="Select which bet to settle:", color=COLORS["green"])
        await interaction.response.send_message(embed=embed, view=SettlePendingView(self.user_id, pending), ephemeral=True)


class SettlePendingView(View):
    def __init__(self, user_id: str, pending_bets: list):
        super().__init__(timeout=120)
        self.user_id = user_id
        self.pending_bets = pending_bets
        
        for i, bet in enumerate(pending_bets[:5]):  # Max 5 buttons
            btn = Button(
                label=f"#{i+1} {bet['description'][:20]}...",
                style=discord.ButtonStyle.secondary,
                custom_id=f"settle_{i}"
            )
            btn.callback = self.make_callback(i)
            self.add_item(btn)
    
    def make_callback(self, index):
        async def callback(interaction: discord.Interaction):
            if str(interaction.user.id) != self.user_id:
                await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
                return
            active_sessions[self.user_id] = {"settling_index": index}
            bet = self.pending_bets[index]
            embed = create_pending_bet_embed(bet, index)
            embed.add_field(name="🏆 Result?", value="Select the outcome:", inline=False)
            await interaction.response.send_message(embed=embed, view=ResultSelectView(self.user_id, settling=True), ephemeral=True)
        return callback


class SportSelectView(View):
    def __init__(self, user_id: str, pending: bool = False):
        super().__init__(timeout=120)
        self.user_id = user_id
        self.pending = pending
        
        for i, sport in enumerate(SPORTS):
            emoji = ["🏀", "🏒", "🏈", "⚾", "⚽", "🥊", "🎾", "🎯"][i]
            btn = Button(label=sport, emoji=emoji, style=discord.ButtonStyle.secondary, custom_id=f"sport_{sport}")
            btn.callback = self.make_callback(sport)
            self.add_item(btn)
    
    def make_callback(self, sport):
        async def callback(interaction: discord.Interaction):
            if str(interaction.user.id) != self.user_id:
                await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
                return
            if self.user_id not in active_sessions:
                active_sessions[self.user_id] = {}
            active_sessions[self.user_id]["sport"] = sport
            active_sessions[self.user_id]["pending"] = self.pending
            
            embed = Embed(
                title=f"{'⏳' if self.pending else '🎲'} {sport}",
                description="Select bet type:",
                color=COLORS["orange"] if self.pending else COLORS["blue"]
            )
            await interaction.response.send_message(embed=embed, view=BetTypeSelectView(self.user_id), ephemeral=True)
        return callback


class BetTypeSelectView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=120)
        self.user_id = user_id
        
        for bet_type in BET_TYPES:
            btn = Button(label=bet_type, style=discord.ButtonStyle.secondary, custom_id=f"type_{bet_type}")
            btn.callback = self.make_callback(bet_type)
            self.add_item(btn)
    
    def make_callback(self, bet_type):
        async def callback(interaction: discord.Interaction):
            if str(interaction.user.id) != self.user_id:
                await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
                return
            if self.user_id in active_sessions:
                session = active_sessions[self.user_id]
                # PRIVATE MODAL for bet details (no public chat messages!)
                modal = BetDataModal(self.user_id, session["sport"], bet_type, pending=session.get("pending", False))
                await interaction.response.send_modal(modal)
        return callback


class ResultSelectView(View):
    def __init__(self, user_id: str, settling: bool = False):
        super().__init__(timeout=120)
        self.user_id = user_id
        self.settling = settling
    
    @ui.button(label="WIN", style=discord.ButtonStyle.green, emoji="✅")
    async def win(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        await self.process_result(interaction, "w")
    
    @ui.button(label="LOSS", style=discord.ButtonStyle.red, emoji="❌")
    async def loss(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        await self.process_result(interaction, "l")
    
    @ui.button(label="PUSH", style=discord.ButtonStyle.gray, emoji="➖")
    async def push(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        await self.process_result(interaction, "p")
    
    async def process_result(self, interaction: discord.Interaction, result: str):
        if self.user_id not in active_sessions:
            await interaction.response.send_message("❌ Session expired. Use /menu to start again.", ephemeral=True)
            return
        
        session = active_sessions[self.user_id]
        data = get_user_data(self.user_id)
        user = data[self.user_id]
        user = reset_if_new_day(user)
        
        # Handle settling a pending bet
        if self.settling and "settling_index" in session:
            index = session["settling_index"]
            pending = user.get("pending_bets", [])
            if index >= len(pending):
                del active_sessions[self.user_id]
                await interaction.response.send_message("❌ Bet not found.", ephemeral=True)
                return
            
            bet = pending.pop(index)
            stake = bet["stake"]
            odds = bet["odds"]
        else:
            # Regular bet
            if user["daily_limit"] is not None and user["bets_today"] >= user["daily_limit"]:
                del active_sessions[self.user_id]
                embed = Embed(title="⚠️ LIMIT REACHED", description="You've hit your daily bet limit!", color=COLORS["red"])
                await interaction.response.send_message(embed=embed, ephemeral=True)
                return
            
            stake = session["stake"]
            odds = session["odds"]
            bet = {
                "sport": session["sport"],
                "bet_type": session["bet_type"],
                "description": session["description"],
                "odds": odds,
                "stake": stake
            }
        
        if result == "w":
            profit = stake * (odds - 1)
            user["wins"] += 1
            user["win_streak"] = user.get("win_streak", 0) + 1
            user["loss_streak"] = 0
            if profit > user.get("best_win", 0):
                user["best_win"] = profit
        elif result == "l":
            profit = -stake
            user["losses"] += 1
            user["loss_streak"] = user.get("loss_streak", 0) + 1
            user["win_streak"] = 0
            if profit < user.get("worst_loss", 0):
                user["worst_loss"] = profit
        else:
            profit = 0
            user["pushes"] = user.get("pushes", 0) + 1
            user["win_streak"] = 0
            user["loss_streak"] = 0
        
        user["total_staked"] += stake
        user["total_profit"] += profit
        user["current_bankroll"] += profit
        user["bets_today"] += 1
        
        bet_entry = {
            "date": datetime.now(TORONTO_TZ).strftime("%Y-%m-%d"),
            "timestamp": datetime.now(TORONTO_TZ).isoformat(timespec="seconds"),
            "sport": bet["sport"],
            "bet_type": bet["bet_type"],
            "description": bet["description"],
            "odds": odds,
            "stake": stake,
            "result": result,
            "profit": profit
        }
        user["bets"].append(bet_entry)
        update_user_data(self.user_id, user)
        
        del active_sessions[self.user_id]
        
        bet_entry["profit"] = profit
        embed = create_bet_recorded_embed(bet_entry, user, result)
        
        # Check for loss alert
        loss_percent = check_loss_alert(user)
        if loss_percent:
            await interaction.response.send_message(embed=embed, ephemeral=True)
            alert_embed = create_loss_alert_embed(loss_percent, user)
            await interaction.followup.send(embed=alert_embed, ephemeral=True)
        else:
            # Check goal progress
            goal_info = get_goal_progress(user)
            if goal_info and goal_info["progress"] >= 100:
                embed.add_field(name="🎉 GOAL REACHED!", value=f"You hit your ${goal_info['goal']:.2f} profit goal!", inline=False)
            await interaction.response.send_message(embed=embed, ephemeral=True)


class LossAlertView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=60)
        self.user_id = user_id
    
    @ui.button(label="10%", style=discord.ButtonStyle.secondary)
    async def alert10(self, interaction: discord.Interaction, button: Button):
        await self.set_alert(interaction, 10)
    
    @ui.button(label="15%", style=discord.ButtonStyle.secondary)
    async def alert15(self, interaction: discord.Interaction, button: Button):
        await self.set_alert(interaction, 15)
    
    @ui.button(label="20%", style=discord.ButtonStyle.secondary)
    async def alert20(self, interaction: discord.Interaction, button: Button):
        await self.set_alert(interaction, 20)
    
    @ui.button(label="25%", style=discord.ButtonStyle.secondary)
    async def alert25(self, interaction: discord.Interaction, button: Button):
        await self.set_alert(interaction, 25)
    
    @ui.button(label="Disable", style=discord.ButtonStyle.red)
    async def disable(self, interaction: discord.Interaction, button: Button):
        await self.set_alert(interaction, 100)
    
    async def set_alert(self, interaction: discord.Interaction, percent: int):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        
        data = get_user_data(self.user_id)
        user = data[self.user_id]
        user["loss_alert_percent"] = percent
        update_user_data(self.user_id, user)
        
        if percent >= 100:
            embed = Embed(title="⚠️ LOSS ALERT", description="Alerts **disabled**", color=COLORS["gray"])
        else:
            embed = Embed(title="⚠️ LOSS ALERT", description=f"Alert set at **{percent}%** daily loss", color=COLORS["green"])
        await interaction.response.send_message(embed=embed, ephemeral=True)


class LimitSelectView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=60)
        self.user_id = user_id
    
    @ui.button(label="3/day", style=discord.ButtonStyle.secondary)
    async def limit3(self, interaction: discord.Interaction, button: Button):
        await self.set_limit(interaction, 3)
    
    @ui.button(label="5/day", style=discord.ButtonStyle.secondary)
    async def limit5(self, interaction: discord.Interaction, button: Button):
        await self.set_limit(interaction, 5)
    
    @ui.button(label="10/day", style=discord.ButtonStyle.secondary)
    async def limit10(self, interaction: discord.Interaction, button: Button):
        await self.set_limit(interaction, 10)
    
    @ui.button(label="No Limit", style=discord.ButtonStyle.red)
    async def no_limit(self, interaction: discord.Interaction, button: Button):
        await self.set_limit(interaction, 0)
    
    async def set_limit(self, interaction: discord.Interaction, limit: int):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        
        data = get_user_data(self.user_id)
        user = data[self.user_id]
        user["daily_limit"] = limit if limit > 0 else None
        update_user_data(self.user_id, user)
        
        if limit <= 0:
            embed = Embed(title="🚧 DAILY LIMIT", description="Limit **disabled**", color=COLORS["gray"])
        else:
            embed = Embed(title="🚧 DAILY LIMIT", description=f"Limit set to **{limit} bets/day**", color=COLORS["green"])
        await interaction.response.send_message(embed=embed, ephemeral=True)


class ConfirmResetView(View):
    def __init__(self, user_id: str):
        super().__init__(timeout=30)
        self.user_id = user_id
    
    @ui.button(label="Yes, Reset Everything", style=discord.ButtonStyle.red)
    async def confirm(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        data = load_data()
        if self.user_id in data:
            data.pop(self.user_id)
            save_data(data)
        embed = Embed(title="♻️ RESET COMPLETE", description="All data has been erased.", color=COLORS["green"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
    
    @ui.button(label="Cancel", style=discord.ButtonStyle.gray)
    async def cancel(self, interaction: discord.Interaction, button: Button):
        if str(interaction.user.id) != self.user_id:
            await interaction.response.send_message("This menu isn't for you!", ephemeral=True)
            return
        embed = Embed(title="✅ Cancelled", description="Your data is safe!", color=COLORS["green"])
        await interaction.response.send_message(embed=embed, ephemeral=True)


# ---------- DISCORD EVENTS ----------

@client.event
async def on_ready():
    print(f"✅ Bot connected as {client.user}")
    try:
        synced = await tree.sync()
        print(f"✅ Synced {len(synced)} command(s)")
    except Exception as e:
        print(f"Error syncing commands: {e}")


@client.event
async def on_guild_join(guild: discord.Guild):
    """Auto-setup when bot joins a server"""
    print(f"Bot joined guild: {guild.name}")
    
    try:
        # Define channels to create
        channels_config = {
            "📊 BANKROLL SYSTEM": [
                ("🎰", "bankroll-bot", "Use /menu to track your bets and manage your bankroll"),
                ("📈", "leaderboard", "View betting stats and compare with friends"),
                ("🎯", "goals-achievements", "Track your profit goals and achievements"),
            ],
            "📢 COMMUNITY": [
                ("💬", "general", "General discussion and chat"),
                ("📍", "picks-sharing", "Share your bets and picks with the community!"),
                ("🏆", "winners", "Share your winning bets!"),
                ("📰", "announcements", "Important updates and announcements"),
            ],
            "ℹ️ INFO": [
                ("👋", "welcome", "Welcome to our betting community!"),
                ("📋", "rules", "Server rules and guidelines"),
                ("🎓", "tutorials", "Guides and how-to tutorials"),
            ]
        }
        
        created_channels = {}
        
        # Create categories and channels
        for category_name, channels in channels_config.items():
            category = await guild.create_category(category_name)
            created_channels[category_name] = {"category": category, "channels": {}}
            
            for emoji, channel_name, description in channels:
                channel = await guild.create_text_channel(
                    f"{emoji}-{channel_name}",
                    category=category,
                    topic=description
                )
                created_channels[category_name]["channels"][channel_name] = channel
        
        # Send welcome message
        welcome_channel = created_channels["ℹ️ INFO"]["channels"]["welcome"]
        welcome_embed = Embed(
            title="🎰 Welcome to Bankroll Manager!",
            description="Your ultimate sports betting tracking community",
            color=COLORS["purple"]
        )
        welcome_embed.add_field(
            name="📊 What is This?",
            value="A Discord bot that helps you track your sports bets, manage your bankroll, set profit goals, and analyze your betting performance.",
            inline=False
        )
        welcome_embed.add_field(
            name="🚀 Quick Start",
            value="1. Go to **#🎰-bankroll-bot**\n2. Type `/start 100` to set your starting bankroll\n3. Type `/menu` to start tracking bets!",
            inline=False
        )
        welcome_embed.add_field(
            name="✨ Features",
            value="✅ Track all your bets\n✅ Pending bets system\n✅ Profit goals & progress bars\n✅ Loss alerts\n✅ Daily summaries\n✅ Excel reports\n✅ Leaderboards\n✅ Weekly analytics",
            inline=False
        )
        welcome_embed.set_footer(text="Good luck! Bet responsibly! 🍀")
        await welcome_channel.send(embed=welcome_embed)
        
        # Send rules
        rules_channel = created_channels["ℹ️ INFO"]["channels"]["rules"]
        rules_embed = Embed(title="📋 Server Rules & Guidelines", color=COLORS["blue"])
        rules_embed.add_field(name="1️⃣ Respect Everyone", value="Be respectful, kind, and supportive.", inline=False)
        rules_embed.add_field(name="2️⃣ Bet Responsibly", value="Only bet what you can afford to lose. Never chase losses.", inline=False)
        rules_embed.add_field(name="3️⃣ No Spam", value="Keep channels clean and organized.", inline=False)
        rules_embed.add_field(name="4️⃣ Stay On Topic", value="Keep discussions relevant to betting and bankroll management.", inline=False)
        rules_embed.add_field(name="5️⃣ Legal Disclaimer", value="Gambling may be illegal in your jurisdiction. Check local laws.", inline=False)
        rules_embed.set_footer(text="Violations may result in warnings or muting.")
        await rules_channel.send(embed=rules_embed)
        
        # Send tutorials
        tutorials_channel = created_channels["ℹ️ INFO"]["channels"]["tutorials"]
        tutorials_embed = Embed(title="🎓 How to Use the Bankroll Bot", color=COLORS["gold"])
        tutorials_embed.add_field(name="Step 1: Initialize Your Bankroll", value="```\n/start 100\n```\nSet your starting amount (e.g., $100)", inline=False)
        tutorials_embed.add_field(name="Step 2: Open the Menu", value="```\n/menu\n```\nClick buttons to track bets, view stats", inline=False)
        tutorials_embed.add_field(name="Step 3: Add a Bet", value="Click **New Bet** → Select sport → Pick type → Enter details", inline=False)
        tutorials_embed.add_field(name="Step 4: Track Results", value="Record the result (WIN/LOSS/PUSH) to update bankroll", inline=False)
        tutorials_embed.add_field(name="Step 5: Use Pending Bets", value="Save bets that aren't settled yet and settle them later!", inline=False)
        tutorials_embed.add_field(name="💡 Pro Tips", value="• Set a **profit goal** to stay motivated\n• Enable **loss alerts** to protect your bankroll\n• Set a **daily bet limit** to prevent tilt\n• Export to **Excel** to analyze your stats", inline=False)
        tutorials_embed.set_footer(text="Questions? Check #general or use /help in #bankroll-bot")
        await tutorials_channel.send(embed=tutorials_embed)
        
        # Send COMPLETE COMMAND LIST
        commands_embed = Embed(title="📖 COMPLETE COMMAND LIST (14 Commands)", description="All commands available to you:", color=COLORS["blue"])
        commands_embed.add_field(name="🎰 GETTING STARTED", value="`/start <amount>` - Set starting bankroll\n`/menu` - Open main menu\n`/help` - Show command guide", inline=False)
        commands_embed.add_field(name="📊 TRACKING YOUR BETS", value="`/menu` - Add bets, settle pending bets, view stats\n`/post-pick` - Share picks in #picks-sharing", inline=False)
        commands_embed.add_field(name="📈 ANALYTICS & PERFORMANCE", value="`/weekly-stats` - This week's record & profit\n`/best-sport` - Highest ROI sport\n`/recovery` - Break-even calculation", inline=False)
        commands_embed.add_field(name="⚙️ SETTINGS & PROTECTION", value="`/daily-cap <percent>` - Max daily loss %\n`/streak-alerts <true/false>` - Get notified\n`/odds-watch <sport> <threshold>` - Track odds", inline=False)
        commands_embed.add_field(name="📋 PLANNING BEFORE YOU BET", value="`/parlay-preview <legs> <odds>` - Calculate parlay\nExample: `/parlay-preview 3 1.5,2.0,1.8`", inline=False)
        commands_embed.add_field(name="🏆 COMMUNITY (EVERYONE SEES)", value="`/leaderboard` - Top 5 bettors\n`/achievements [user]` - View badges", inline=False)
        commands_embed.set_footer(text="Type any command to get started! All replies are private (only you see them)")
        await tutorials_channel.send(embed=commands_embed)
        
        # Send DETAILED COMMAND GUIDE
        detail_embed = Embed(title="🎯 HOW EACH COMMAND WORKS (DETAILED)", color=COLORS["purple"])
        
        detail_embed.add_field(
            name="1️⃣ /start <amount>",
            value="**What:** Initialize your bankroll\n**How:** Type `/start 100` to start with $100\n**When:** Use this FIRST before anything else\n**Example:** `/start 500`",
            inline=False
        )
        
        detail_embed.add_field(
            name="2️⃣ /menu",
            value="**What:** Main hub with 5 buttons\n**How:** Type `/menu` and click buttons:\n• **New Bet** - Add a bet\n• **Pending Bets** - Settle unsettled bets\n• **Summary** - View overall stats\n• **Daily Report** - Today's results\n• **Export Excel** - Download betting history",
            inline=False
        )
        
        detail_embed.add_field(
            name="3️⃣ /post-pick <sport> <pick> <odds> <confidence>",
            value="**What:** Share your pick in #picks-sharing\n**How:** `/post-pick NBA \"Lakers -5.5\" 1.85 High`\n**Confidence:** Low, Medium, or High\n**Result:** Posted to #picks-sharing with reactions (✅/❌)",
            inline=False
        )
        
        detail_embed.add_field(
            name="4️⃣ /daily-cap <percent>",
            value="**What:** Set maximum daily loss\n**How:** `/daily-cap 5` = can't lose more than 5% of bankroll per day\n**Example:** If bankroll is $1000, max loss is $50/day\n**Protection:** Prevents over-betting when you're down",
            inline=False
        )
        
        detail_embed.add_field(
            name="5️⃣ /recovery",
            value="**What:** Calculate break-even amount\n**How:** Type `/recovery`\n**Shows:** How much profit you need to get back to $0\n**Tip:** Helps you stay realistic about recovery goals",
            inline=False
        )
        
        detail_embed.add_field(
            name="6️⃣ /best-sport",
            value="**What:** Shows your highest ROI sport\n**How:** Type `/best-sport`\n**Result:** See which sport makes you the most money %\n**Tip:** Focus on your best-performing sports!",
            inline=False
        )
        
        detail_embed.add_field(
            name="7️⃣ /streak-alerts <true/false>",
            value="**What:** Get notifications at 5+ streaks\n**How:** `/streak-alerts true` to enable\n**Alerts:** You'll get notified on big win/loss streaks\n**Use:** Helps you avoid tilting after bad streaks",
            inline=False
        )
        
        detail_embed.add_field(
            name="8️⃣ /parlay-preview <legs> <odds>",
            value="**What:** Calculate parlay payout\n**How:** `/parlay-preview 3 1.5,2.0,1.8`\n**Shows:** Total odds, profit on $100 stake\n**Tip:** Preview before betting real money!",
            inline=False
        )
        
        detail_embed.add_field(
            name="9️⃣ /weekly-stats",
            value="**What:** This week's performance\n**How:** Type `/weekly-stats`\n**Shows:** Record, profit, ROI, # of bets\n**Useful:** Track weekly trends",
            inline=False
        )
        
        detail_embed.add_field(
            name="🔟 /odds-watch <sport> <threshold>",
            value="**What:** Track important odds\n**How:** `/odds-watch NBA 2.0`\n**Stores:** Up to 5 watched odds\n**Use:** Remember odds you want to track",
            inline=False
        )
        
        detail_embed.add_field(
            name="1️⃣1️⃣ /leaderboard",
            value="**What:** Top 5 bettors by profit\n**How:** Type `/leaderboard`\n**Shows:** Rankings, profit, win rate\n**Who sees:** EVERYONE on the server",
            inline=False
        )
        
        detail_embed.add_field(
            name="1️⃣2️⃣ /achievements [user]",
            value="**What:** View badges/achievements\n**How:** `/achievements` = your badges\n**How:** `/achievements @username` = see someone else's\n**Badges:** Unlock by hitting milestones (profit, streaks, etc)",
            inline=False
        )
        
        detail_embed.add_field(
            name="1️⃣3️⃣ /help",
            value="**What:** Show all commands with descriptions\n**How:** Type `/help`\n**Shows:** Organized list of all 14 commands",
            inline=False
        )
        
        detail_embed.set_footer(text="All commands are PRIVATE (only you see results) except /leaderboard and /achievements")
        await tutorials_channel.send(embed=detail_embed)
        
        print(f"✅ Auto-setup complete for {guild.name}")
        
    except Exception as e:
        print(f"Error during auto-setup: {e}")


@client.event
async def on_message(message: discord.Message):
    if message.author == client.user:
        return

    content = message.content.strip()
    lower = content.lower()
    user_id = str(message.author.id)

    try:
        # Check if user is in a session
        if user_id in active_sessions:
            session = active_sessions[user_id]
            step = session.get("step")
            
            if step == "set_goal":
                try:
                    goal = float(content.replace("$", "").replace(",", ""))
                    if goal <= 0:
                        embed = Embed(title="❌ Invalid", description="Goal must be positive.", color=COLORS["red"])
                        await message.channel.send(embed=embed, delete_after=5)
                        return
                    
                    data = get_user_data(user_id)
                    user = data[user_id]
                    user["profit_goal"] = goal
                    update_user_data(user_id, user)
                    del active_sessions[user_id]
                    
                    progress = get_goal_progress(user)
                    bar = create_progress_bar(progress["progress"])
                    
                    embed = Embed(title="🎯 GOAL SET!", color=COLORS["gold"])
                    embed.add_field(name="Target", value=f"**${goal:.2f}** profit", inline=True)
                    embed.add_field(name="Current", value=f"**${user['total_profit']:.2f}**", inline=True)
                    embed.add_field(name="Progress", value=f"{bar} **{progress['progress']:.1f}%**", inline=False)
                    await message.channel.send(embed=embed, delete_after=30)
                except ValueError:
                    embed = Embed(title="❌ Invalid", description="Enter a number like `500`", color=COLORS["red"])
                    await message.channel.send(embed=embed, delete_after=5)
                return
            
            if step == "description":
                session["description"] = content
                session["step"] = "odds"
                embed = Embed(
                    title="📊 ODDS",
                    description="Enter the odds (decimal):\n\nExample: `1.85` or `2.10`",
                    color=COLORS["blue"]
                )
                await message.channel.send(embed=embed, delete_after=30)
                return
            
            elif step == "odds":
                try:
                    odds = float(content.replace(",", "."))
                    if odds <= 1:
                        embed = Embed(title="❌ Invalid", description="Odds must be greater than 1.", color=COLORS["red"])
                        await message.channel.send(embed=embed, delete_after=5)
                        return
                    session["odds"] = odds
                    session["step"] = "stake"
                    embed = Embed(
                        title="💵 STAKE",
                        description="Enter your stake:\n\nExample: `25` or `50.00`",
                        color=COLORS["blue"]
                    )
                    await message.channel.send(embed=embed, delete_after=30)
                except ValueError:
                    embed = Embed(title="❌ Invalid", description="Use decimal like `1.85`", color=COLORS["red"])
                    await message.channel.send(embed=embed, delete_after=5)
                return
            
            elif step == "stake":
                try:
                    stake = float(content.replace("$", "").replace(",", "."))
                    if stake <= 0:
                        embed = Embed(title="❌ Invalid", description="Stake must be positive.", color=COLORS["red"])
                        await message.channel.send(embed=embed, delete_after=5)
                        return
                    session["stake"] = stake
                    
                    # Check if this is a pending bet
                    if session.get("pending"):
                        # Save as pending
                        data = get_user_data(user_id)
                        user = data[user_id]
                        
                        pending_bet = {
                            "date": datetime.now(TORONTO_TZ).strftime("%Y-%m-%d"),
                            "sport": session["sport"],
                            "bet_type": session["bet_type"],
                            "description": session["description"],
                            "odds": session["odds"],
                            "stake": stake
                        }
                        user["pending_bets"].append(pending_bet)
                        update_user_data(user_id, user)
                        del active_sessions[user_id]
                        
                        potential = stake * (session["odds"] - 1)
                        sport_emoji = SPORT_EMOJIS.get(session["sport"], "🎯")
                        
                        embed = Embed(title="⏳ PENDING BET ADDED", color=COLORS["orange"])
                        embed.add_field(
                            name="📋 Details",
                            value=f"{sport_emoji} **{session['sport']}** | {session['bet_type']}\n"
                                  f"📝 {session['description']}\n"
                                  f"📊 Odds: **{session['odds']}**\n"
                                  f"💵 Stake: **${stake:.2f}**",
                            inline=False
                        )
                        embed.add_field(name="💰 Potential Win", value=f"**+${potential:.2f}**", inline=True)
                        embed.set_footer(text="Use /menu → Pending Bets → Settle when it's done!")
                        await message.channel.send(embed=embed, delete_after=60)
                    else:
                        session["step"] = "result"
                        embed = Embed(
                            title="🏆 RESULT",
                            description="Select the outcome:",
                            color=COLORS["blue"]
                        )
                        await message.channel.send(embed=embed, view=ResultSelectView(user_id), delete_after=60)
                except ValueError:
                    embed = Embed(title="❌ Invalid", description="Use numbers like `25`", color=COLORS["red"])
                    await message.channel.send(embed=embed, delete_after=5)
                return

    except Exception as e:
        print("Error:", e)


# ---------- SLASH COMMANDS ----------

@tree.command(name="menu", description="Open the bankroll management menu")
async def menu_command(interaction: discord.Interaction):
    user_id = str(interaction.user.id)
    embed = create_menu_embed(user_id)
    await interaction.response.send_message(embed=embed, view=MainMenuView(user_id), ephemeral=True)


@tree.command(name="start", description="Initialize your bankroll")
async def start_command(interaction: discord.Interaction, amount: float):
    user_id = str(interaction.user.id)
    
    if amount <= 0:
        embed = Embed(title="❌ Invalid Amount", description="Bankroll must be positive!", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    data = load_data()
    data[user_id] = {
        "starting_bankroll": amount,
        "current_bankroll": amount,
        "total_staked": 0.0,
        "total_profit": 0.0,
        "wins": 0,
        "losses": 0,
        "pushes": 0,
        "bets": [],
        "pending_bets": [],
        "daily_limit": None,
        "bets_today": 0,
        "last_date": str(date.today()),
        "profit_goal": None,
        "loss_alert_percent": 20,
        "daily_starting_bankroll": amount,
        "last_summary_date": None,
        "win_streak": 0,
        "loss_streak": 0,
        "best_win": 0.0,
        "worst_loss": 0.0,
        "username": interaction.user.name
    }
    save_data(data)
    
    embed = Embed(title="🚀 BANKROLL SET!", color=COLORS["green"])
    embed.add_field(name="💰 Starting Balance", value=f"**${amount:.2f}**", inline=False)
    embed.add_field(name="📱 Next Step", value="Use `/menu` to access all features!", inline=False)
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="setup", description="Setup the server with channels and welcome messages (Admin only)")
async def setup_command(interaction: discord.Interaction):
    print("Setup command called!")
    
    # Check if user is admin
    if not interaction.user.guild_permissions.administrator:
        embed = Embed(title="❌ Admin Only", description="Only server admins can use this command!", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        print("User is not admin")
        return
    
    print("User is admin, deferring...")
    await interaction.response.defer(ephemeral=True)
    
    guild = interaction.guild
    print(f"Guild: {guild.name}")
    
    try:
        # Define channels to create
        channels_config = {
            "📊 BANKROLL SYSTEM": [
                ("🎰", "bankroll-bot", "Use /menu to track your bets and manage your bankroll"),
                ("📈", "leaderboard", "View betting stats and compare with friends"),
                ("🎯", "goals-achievements", "Track your profit goals and achievements"),
            ],
            "📢 COMMUNITY": [
                ("💬", "general", "General discussion and chat"),
                ("📍", "picks-sharing", "Share your bets and picks with the community!"),
                ("🏆", "winners", "Share your winning bets!"),
                ("📰", "announcements", "Important updates and announcements"),
            ],
            "ℹ️ INFO": [
                ("👋", "welcome", "Welcome to our betting community!"),
                ("📋", "rules", "Server rules and guidelines"),
                ("🎓", "tutorials", "Guides and how-to tutorials"),
            ]
        }
        
        created_channels = {}
        
        # Create categories and channels
        for category_name, channels in channels_config.items():
            # Create category
            category = await guild.create_category(category_name)
            created_channels[category_name] = {"category": category, "channels": {}}
            
            # Create channels in category
            for emoji, channel_name, description in channels:
                channel = await guild.create_text_channel(
                    f"{emoji}-{channel_name}",
                    category=category,
                    topic=description
                )
                created_channels[category_name]["channels"][channel_name] = channel
        
        # Send welcome message to welcome channel
        welcome_channel = created_channels["ℹ️ INFO"]["channels"]["welcome"]
        
        welcome_embed = Embed(
            title="🎰 Welcome to Bankroll Manager!",
            description="Your ultimate sports betting tracking community",
            color=COLORS["purple"]
        )
        
        welcome_embed.add_field(
            name="📊 What is This?",
            value="A Discord bot that helps you track your sports bets, manage your bankroll, set profit goals, and analyze your betting performance.",
            inline=False
        )
        
        welcome_embed.add_field(
            name="🚀 Quick Start",
            value="1. Go to **#🎰-bankroll-bot**\n2. Type `/start 100` to set your starting bankroll\n3. Type `/menu` to start tracking bets!",
            inline=False
        )
        
        welcome_embed.add_field(
            name="✨ Features",
            value="✅ Track all your bets\n✅ Pending bets system\n✅ Profit goals & progress bars\n✅ Loss alerts\n✅ Daily summaries\n✅ Excel reports\n✅ Leaderboards",
            inline=False
        )
        
        welcome_embed.add_field(
            name="📍 Server Channels",
            value="**Bankroll System:** Track bets, view leaderboards, set goals\n**Community:** Chat, share wins, get announcements\n**Info:** Rules, tutorials, this welcome message",
            inline=False
        )
        
        welcome_embed.set_footer(text="Good luck! Bet responsibly! 🍀")
        welcome_embed.set_thumbnail(url="https://cdn.discordapp.com/emojis/1234567890.png")
        
        await welcome_channel.send(embed=welcome_embed)
        
        # Send rules message
        rules_channel = created_channels["ℹ️ INFO"]["channels"]["rules"]
        
        rules_embed = Embed(
            title="📋 Server Rules & Guidelines",
            color=COLORS["blue"]
        )
        
        rules_embed.add_field(
            name="1️⃣ Respect Everyone",
            value="Be respectful, kind, and supportive. No harassment or discrimination.",
            inline=False
        )
        
        rules_embed.add_field(
            name="2️⃣ Bet Responsibly",
            value="Only bet what you can afford to lose. Never chase losses. Use stop-loss limits.",
            inline=False
        )
        
        rules_embed.add_field(
            name="3️⃣ No Spam",
            value="Don't spam messages, emojis, or links. Keep channels clean and organized.",
            inline=False
        )
        
        rules_embed.add_field(
            name="4️⃣ Stay On Topic",
            value="Keep discussions relevant to betting and bankroll management.",
            inline=False
        )
        
        rules_embed.add_field(
            name="5️⃣ Legal Disclaimer",
            value="Gambling may be illegal in your jurisdiction. Check local laws before betting.",
            inline=False
        )
        
        rules_embed.set_footer(text="Violations may result in warnings or muting.")
        
        await rules_channel.send(embed=rules_embed)
        
        # Send tutorials message
        tutorials_channel = created_channels["ℹ️ INFO"]["channels"]["tutorials"]
        
        tutorials_embed = Embed(
            title="🎓 How to Use the Bankroll Bot",
            color=COLORS["gold"]
        )
        
        tutorials_embed.add_field(
            name="Step 1: Initialize Your Bankroll",
            value="```\n/start 100\n```\nSet your starting amount (e.g., $100)",
            inline=False
        )
        
        tutorials_embed.add_field(
            name="Step 2: Open the Menu",
            value="```\n/menu\n```\nClick buttons to track bets, view stats, and manage settings",
            inline=False
        )
        
        tutorials_embed.add_field(
            name="Step 3: Add a Bet",
            value="Click **New Bet** → Select sport → Pick bet type → Enter pick, odds, and stake",
            inline=False
        )
        
        tutorials_embed.add_field(
            name="Step 4: Track Results",
            value="After the bet settles, record the result (WIN/LOSS/PUSH) to update your bankroll",
            inline=False
        )
        
        tutorials_embed.add_field(
            name="Step 5: Use Pending Bets",
            value="Don't know the result yet? Use **Pending Bets** to save bets and settle them later!",
            inline=False
        )
        
        tutorials_embed.add_field(
            name="💡 Pro Tips",
            value="• Set a **profit goal** to stay motivated\n• Enable **loss alerts** to protect your bankroll\n• Set a **daily bet limit** to prevent tilt\n• Export to **Excel** to analyze your stats",
            inline=False
        )
        
        await tutorials_channel.send(embed=tutorials_embed)
        
        # Send ALL COMMANDS guide
        commands_embed = Embed(
            title="📖 COMPLETE COMMAND LIST (14 Commands)",
            description="All commands available to you:",
            color=COLORS["blue"]
        )
        
        commands_embed.add_field(
            name="🎰 GETTING STARTED",
            value="`/start <amount>` - Set your starting bankroll\n`/menu` - Open main menu with buttons\n`/help` - Show this guide",
            inline=False
        )
        
        commands_embed.add_field(
            name="📊 TRACKING YOUR BETS",
            value="`/menu` - Add bets, settle pending bets, view daily/overall stats\n`/post-pick <sport> <pick> <odds> <confidence>` - Share picks in #picks-sharing",
            inline=False
        )
        
        commands_embed.add_field(
            name="📈 ANALYTICS & PERFORMANCE",
            value="`/weekly-stats` - See this week's record, profit, ROI\n`/best-sport` - Find your highest ROI sport\n`/recovery` - Calculate how much you need to win to break even",
            inline=False
        )
        
        commands_embed.add_field(
            name="⚙️ SETTINGS & PROTECTION",
            value="`/daily-cap <percent>` - Set max daily loss (e.g., 5%)\n`/streak-alerts <true/false>` - Get notified at 5+ win/loss streaks\n`/odds-watch <sport> <threshold>` - Track important odds",
            inline=False
        )
        
        commands_embed.add_field(
            name="📋 PLANNING BEFORE YOU BET",
            value="`/parlay-preview <legs> <odds>` - Calculate parlay potential before placing\nExample: `/parlay-preview 3 1.5,2.0,1.8`",
            inline=False
        )
        
        commands_embed.add_field(
            name="🏆 COMMUNITY (EVERYONE SEES)",
            value="`/leaderboard` - Top 5 bettors by profit\n`/achievements [user]` - View your badges or someone else's",
            inline=False
        )
        
        commands_embed.add_field(
            name="🛠️ ADMIN ONLY",
            value="`/setup` - Create all 9 channels (run once!)",
            inline=False
        )
        
        commands_embed.set_footer(text="Type any command to get started! All replies are private (only you see them)")
        
        await tutorials_channel.send(embed=commands_embed)
        
        # Send leaderboard intro
        leaderboard_channel = created_channels["📊 BANKROLL SYSTEM"]["channels"]["leaderboard"]
        
        leaderboard_embed = Embed(
            title="🏆 Leaderboard",
            description="Your betting stats at a glance",
            color=COLORS["green"]
        )
        
        leaderboard_embed.add_field(
            name="Coming Soon",
            value="Use `/menu` → Summary to see your personal stats!\n\nTip: Export to Excel for detailed analysis.",
            inline=False
        )
        
        leaderboard_embed.set_footer(text="Track your progress and compete with friends!")
        
        await leaderboard_channel.send(embed=leaderboard_embed)
        
        # Success message
        success_embed = Embed(
            title="✅ Server Setup Complete!",
            description="Your Discord server is now ready for bankroll tracking!",
            color=COLORS["green"]
        )
        
        success_embed.add_field(
            name="📍 Created Channels",
            value="✅ 📊 Bankroll System (3 channels)\n✅ 📢 Community (3 channels)\n✅ ℹ️ Info (3 channels)",
            inline=False
        )
        
        success_embed.add_field(
            name="🚀 Next Steps",
            value="1. Go to **#🎰-bankroll-bot**\n2. Users can type `/start <amount>` to initialize\n3. Type `/menu` to start tracking!",
            inline=False
        )
        
        success_embed.add_field(
            name="📝 Customization",
            value="You can rename channels, reorder them, and set permissions as needed!",
            inline=False
        )
        
        success_embed.set_footer(text="Happy betting! 🎰")
        
        await interaction.followup.send(embed=success_embed, ephemeral=True)
        
    except Exception as e:
        print(f"Setup error: {type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        error_embed = Embed(
            title="❌ Setup Failed",
            description=f"Error: {str(e)[:100]}",
            color=COLORS["red"]
        )
        try:
            await interaction.followup.send(embed=error_embed, ephemeral=True)
        except:
            print("Failed to send error message")


@tree.command(name="leaderboard", description="View the top 5 betters by profit")
async def leaderboard_command(interaction: discord.Interaction):
    lb = get_leaderboard()
    
    if not lb:
        embed = Embed(title="📈 LEADERBOARD", description="No bettors yet!", color=COLORS["gray"])
        await interaction.response.send_message(embed=embed)
        return
    
    embed = Embed(title="🏆 TOP 5 BETTORS 🏆", color=COLORS["gold"])
    
    medals = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣"]
    
    for i, user in enumerate(lb):
        medal = medals[i]
        value = f"**+${user['profit']:.2f}**\nRecord: {user['record']} | {user['wr']:.1f}% 🎯"
        embed.add_field(name=f"{medal} {user['name']}", value=value, inline=False)
    
    embed.set_footer(text="🚀 Keep grinding! Your profit goal is waiting!")
    await interaction.response.send_message(embed=embed)


@tree.command(name="achievements", description="View achievements and badges")
async def achievements_command(interaction: discord.Interaction, user: discord.User = None):
    if user is None:
        user = interaction.user
    
    user_id = str(user.id)
    data = get_user_data(user_id)
    user_data = data[user_id]
    
    achievements = get_achievements(user_data)
    
    embed = Embed(title=f"🎖️ {user.name}'S ACHIEVEMENTS", color=COLORS["purple"])
    
    if not achievements:
        embed.add_field(
            name="No Achievements Yet",
            value="Start betting to unlock achievements!",
            inline=False
        )
    else:
        for emoji, name, desc in achievements:
            embed.add_field(name=f"{emoji} {name}", value=desc, inline=True)
    
    embed.add_field(
        name="📊 Stats",
        value=f"**{user_data['wins']}W - {user_data['losses']}L** | **+${user_data['total_profit']:+.2f}** | 🔥 {user_data.get('win_streak', 0)} streak",
        inline=False
    )
    
    await interaction.response.send_message(embed=embed)


@tree.command(name="post-pick", description="Share your bet or pick with the community")
async def post_pick_command(
    interaction: discord.Interaction, 
    sport: str, 
    pick: str, 
    odds: float, 
    confidence: str = "Medium"
):
    embed = Embed(
        title="📍 NEW PICK POSTED!",
        color=COLORS["blue"]
    )
    
    # Confidence emoji
    confidence_map = {
        "low": ("🔵", Color.blue()),
        "medium": ("🟡", Color.gold()),
        "high": ("🟢", Color.green()),
    }
    emoji, color = confidence_map.get(confidence.lower(), ("🟡", Color.gold()))
    embed.color = color
    
    embed.add_field(
        name=f"{emoji} {interaction.user.name}'s Pick",
        value=f"**{sport.upper()}**",
        inline=False
    )
    
    embed.add_field(
        name="📝 Pick",
        value=f"**{pick}**",
        inline=True
    )
    
    embed.add_field(
        name="📊 Odds",
        value=f"**{odds}**",
        inline=True
    )
    
    embed.add_field(
        name="💪 Confidence",
        value=f"**{confidence}**",
        inline=True
    )
    
    embed.set_footer(text="💬 React with ✅ if you like this pick!")
    
    # Send to picks-sharing channel if it exists
    try:
        channel = discord.utils.get(interaction.guild.text_channels, name="📍-picks-sharing")
        if channel:
            msg = await channel.send(embed=embed)
            await msg.add_reaction("✅")
            await msg.add_reaction("❌")
            await interaction.response.send_message(
                f"✅ Your pick was posted to #picks-sharing!", 
                ephemeral=True
            )
        else:
            await interaction.response.send_message(embed=embed, ephemeral=True)
    except:
        await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="daily-cap", description="Set maximum daily loss cap (% of bankroll)")
async def daily_cap_command(interaction: discord.Interaction, percent: float):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    if percent <= 0 or percent > 50:
        embed = Embed(title="❌ Invalid", description="Percent must be 1-50%", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    user["daily_cap_percent"] = percent
    update_user_data(user_id, user)
    
    cap_amount = user["starting_bankroll"] * (percent / 100)
    embed = Embed(
        title="💰 DAILY CAP SET",
        description=f"Max loss per day: **{percent}%** (${cap_amount:.2f})",
        color=COLORS["green"]
    )
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="recovery", description="Calculate how much you need to win to break even")
async def recovery_command(interaction: discord.Interaction):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    current_profit = user["total_profit"]
    
    if current_profit >= 0:
        embed = Embed(
            title="✅ YOU'RE UP!",
            description=f"Profit: **+${current_profit:.2f}**\n\nNo recovery needed! 🎉",
            color=COLORS["green"]
        )
    else:
        # Amount needed to break even
        needed = abs(current_profit)
        avg_odds = (user["total_staked"] / len(user["bets"])) if user["bets"] else 1.5
        implied_odds = 1 + (needed / user.get("avg_stake", 50))
        
        embed = Embed(
            title="📊 RECOVERY CALCULATOR",
            color=COLORS["orange"]
        )
        embed.add_field(
            name="📉 Current Position",
            value=f"**${current_profit:.2f}** down",
            inline=False
        )
        embed.add_field(
            name="💰 Needed to Break Even",
            value=f"**+${needed:.2f}** profit",
            inline=True
        )
        embed.add_field(
            name="🎲 At 1.5 odds",
            value=f"Need: **${needed/0.5:.2f}** stake",
            inline=True
        )
        embed.add_field(
            name="💡 Tip",
            value="Focus on quality picks, not chasing losses!",
            inline=False
        )
    
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="best-sport", description="View your best performing sport (ROI %)")
async def best_sport_command(interaction: discord.Interaction):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    if not user["bets"]:
        embed = Embed(title="📊 NO BETS YET", description="Place some bets first!", color=COLORS["gray"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    # Calculate per-sport ROI
    sport_stats = {}
    for bet in user["bets"]:
        sport = bet.get("sport", "Other")
        if sport not in sport_stats:
            sport_stats[sport] = {"profit": 0, "staked": 0, "wins": 0, "bets": 0}
        sport_stats[sport]["profit"] += bet.get("profit", 0)
        sport_stats[sport]["staked"] += bet.get("stake", 0)
        sport_stats[sport]["bets"] += 1
        if bet.get("result") == "w":
            sport_stats[sport]["wins"] += 1
    
    # Find best sport by ROI
    best = None
    best_roi = -999
    for sport, stats in sport_stats.items():
        roi = (stats["profit"] / stats["staked"] * 100) if stats["staked"] > 0 else 0
        if roi > best_roi:
            best_roi = roi
            best = (sport, stats, roi)
    
    if best:
        sport, stats, roi = best
        wr = (stats["wins"] / stats["bets"] * 100) if stats["bets"] > 0 else 0
        emoji = SPORT_EMOJIS.get(sport, "🎯")
        
        embed = Embed(
            title=f"🔥 YOUR BEST SPORT: {emoji} {sport}",
            color=COLORS["green"] if roi > 0 else COLORS["red"]
        )
        embed.add_field(
            name="📊 ROI",
            value=f"**{roi:+.2f}%**",
            inline=True
        )
        embed.add_field(
            name="💰 Profit",
            value=f"**${stats['profit']:+.2f}**",
            inline=True
        )
        embed.add_field(
            name="🏆 Record",
            value=f"**{stats['wins']}-{stats['bets']-stats['wins']}** ({wr:.1f}%)",
            inline=True
        )
        
        embed.set_footer(text="Focus on your strengths! 🎯")
        await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="streak-alerts", description="Enable/disable streak notifications")
async def streak_alerts_command(interaction: discord.Interaction, enabled: bool):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    user["streak_alerts"] = enabled
    update_user_data(user_id, user)
    
    status = "🔔 **ENABLED**" if enabled else "🔕 **DISABLED**"
    embed = Embed(
        title="🔥 STREAK ALERTS",
        description=f"Notifications: {status}\n\nYou'll get alerts at 5+ win/loss streaks!",
        color=COLORS["green"] if enabled else COLORS["gray"]
    )
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="parlay-preview", description="Preview parlay potential before placing")
async def parlay_preview_command(interaction: discord.Interaction, legs: str, odds_list: str):
    try:
        odds = [float(x.strip()) for x in odds_list.split(",")]
        total_odds = 1
        for odd in odds:
            total_odds *= odd
        
        stake = 100  # Preview with $100
        win = stake * (total_odds - 1)
        
        embed = Embed(
            title="📋 PARLAY PREVIEW",
            color=COLORS["blue"]
        )
        embed.add_field(
            name="🎯 Legs",
            value=f"**{legs}**",
            inline=False
        )
        embed.add_field(
            name="📊 Implied Odds",
            value=f"**{total_odds:.2f}**",
            inline=True
        )
        embed.add_field(
            name="💰 @$100 stake",
            value=f"**+${win:.2f}** profit",
            inline=True
        )
        embed.add_field(
            name="⚠️ Risk",
            value=f"Lose: **$100** | Win: **${stake + win:.2f}**",
            inline=False
        )
        
        await interaction.response.send_message(embed=embed, ephemeral=True)
    except:
        embed = Embed(
            title="❌ Invalid Format",
            description="Usage: `/parlay-preview legs:3 odds_list:1.5,2.0,1.8`",
            color=COLORS["red"]
        )
        await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="weekly-stats", description="View your performance this week")
async def weekly_stats_command(interaction: discord.Interaction):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    if not user["bets"]:
        embed = Embed(title="📊 NO BETS YET", description="Place some bets first!", color=COLORS["gray"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    # Get this week's bets
    today = date.today()
    week_start = today - timedelta(days=today.weekday())
    week_bets = []
    for bet in user["bets"]:
        try:
            bet_date = datetime.strptime(bet.get("date", ""), "%Y-%m-%d").date()
            if bet_date >= week_start:
                week_bets.append(bet)
        except:
            pass
    
    if not week_bets:
        embed = Embed(
            title="📊 THIS WEEK",
            description="No bets this week yet!",
            color=COLORS["gray"]
        )
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    # Calculate stats
    wins = sum(1 for b in week_bets if b.get("result") == "w")
    losses = sum(1 for b in week_bets if b.get("result") == "l")
    profit = sum(b.get("profit", 0) for b in week_bets)
    staked = sum(b.get("stake", 0) for b in week_bets)
    roi = (profit / staked * 100) if staked > 0 else 0
    wr = (wins / len(week_bets) * 100) if week_bets else 0
    
    color = COLORS["green"] if profit >= 0 else COLORS["red"]
    embed = Embed(
        title="📊 THIS WEEK'S PERFORMANCE",
        color=color
    )
    
    embed.add_field(
        name="🏆 Record",
        value=f"**{wins}-{losses}** ({wr:.1f}% WR)",
        inline=True
    )
    embed.add_field(
        name="💰 Profit",
        value=f"**${profit:+.2f}**",
        inline=True
    )
    embed.add_field(
        name="📈 ROI",
        value=f"**{roi:+.2f}%**",
        inline=True
    )
    embed.add_field(
        name="📌 Bets",
        value=f"**{len(week_bets)}** total",
        inline=True
    )
    embed.add_field(
        name="💵 Staked",
        value=f"**${staked:.2f}**",
        inline=True
    )
    
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="odds-watch", description="Track important odds thresholds")
async def odds_watch_command(interaction: discord.Interaction, sport: str, threshold: float):
    user_id = str(interaction.user.id)
    data = get_user_data(user_id)
    user = data[user_id]
    
    if "odds_watch" not in user:
        user["odds_watch"] = []
    
    watch_item = {
        "sport": sport,
        "threshold": threshold,
        "added": datetime.now(TORONTO_TZ).isoformat()
    }
    user["odds_watch"].append(watch_item)
    
    # Keep only last 5
    user["odds_watch"] = user["odds_watch"][-5:]
    update_user_data(user_id, user)
    
    embed = Embed(
        title="👀 ODDS WATCH",
        description=f"Tracking **{sport}** at **{threshold}** odds",
        color=COLORS["blue"]
    )
    embed.add_field(
        name="📌 You'll be reminded",
        value="when odds hit this threshold!",
        inline=False
    )
    embed.add_field(
        name="📍 Watched",
        value=f"**{len(user['odds_watch'])}/5** slots",
        inline=False
    )
    
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="help", description="Get help with the bot")
async def help_command(interaction: discord.Interaction):
    embed = Embed(
        title="🤖 BANKROLL BOT - COMPLETE COMMAND LIST",
        description="All 14 commands available to you:",
        color=COLORS["purple"]
    )
    
    embed.add_field(
        name="🎰 START HERE",
        value="`/start <amount>` - Initialize your bankroll\n`/menu` - Open main menu with buttons",
        inline=False
    )
    
    embed.add_field(
        name="📊 TRACKING & ANALYTICS",
        value="`/menu` - Record bets, view pending bets\n`/weekly-stats` - This week's performance\n`/best-sport` - Your highest ROI sport\n`/recovery` - How much to break even",
        inline=False
    )
    
    embed.add_field(
        name="⚙️ SETTINGS & LIMITS",
        value="`/daily-cap <percent>` - Max daily loss %\n`/streak-alerts <true/false>` - Get notifications\n`/odds-watch <sport> <threshold>` - Track odds",
        inline=False
    )
    
    embed.add_field(
        name="📋 PLANNING & PREVIEW",
        value="`/parlay-preview <legs> <odds>` - Calculate parlay before placing",
        inline=False
    )
    
    embed.add_field(
        name="📢 COMMUNITY (PUBLIC)",
        value="`/leaderboard` - Top 5 bettors by profit\n`/achievements [user]` - View badges (yours or others)\n`/post-pick <sport> <pick> <odds> <confidence>` - Share picks",
        inline=False
    )
    
    embed.add_field(
        name="📚 ADMIN",
        value="`/setup` - Create all channels (admin only)\n`/help` - This message",
        inline=False
    )
    
    embed.set_footer(text="💡 Tip: Start with /start then /menu!")
    await interaction.response.send_message(embed=embed, ephemeral=True)


@tree.command(name="setup-roles", description="Setup Discord roles (Admin only)")
async def setup_roles_command(interaction: discord.Interaction):
    if not interaction.user.guild_permissions.administrator:
        embed = Embed(title="❌ Admin Only", description="Only admins can setup roles!", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    
    await interaction.response.defer(ephemeral=True)
    guild = interaction.guild
    
    try:
        # Define roles to create (with positions for sidebar sections)
        roles_config = [
            ("🔱 Godfather", discord.Color.from_rgb(0, 0, 0)),  # Black
            ("👨‍💼 CEO", discord.Color.from_rgb(255, 0, 0)),  # Red
            ("🚀 Moon Rocket", discord.Color.from_rgb(128, 0, 128)),  # Purple
            ("🏆 Legend", discord.Color.from_rgb(255, 215, 0)),  # Gold
            ("💎 Diamond", discord.Color.from_rgb(0, 191, 255)),  # Light blue
            ("⚡ Elite", discord.Color.from_rgb(255, 255, 0)),  # Yellow
            ("🎯 Sharpshooter", discord.Color.from_rgb(0, 128, 0)),  # Green
            ("💰 Whale", discord.Color.from_rgb(0, 0, 255)),  # Blue
            ("⭐ VIP", discord.Color.gold()),  # Gold
            ("👑 Admin", discord.Color.red()),  # Red
            ("💰 Member", discord.Color.blue()),  # Blue
            ("🤖 Bot", discord.Color.purple())  # Purple
        ]
        
        created_roles = []
        
        # Create roles
        for role_name, color in roles_config:
            # Check if role already exists
            existing_role = discord.utils.find(lambda r: r.name == role_name, guild.roles)
            if existing_role:
                created_roles.append(f"✅ {role_name}")
            else:
                new_role = await guild.create_role(name=role_name, color=color)
                created_roles.append(f"✅ {role_name}")
        
        # Assign multiple cool roles to user (CEO, Moon Rocket, Legend, Godfather, Admin)
        cool_roles = ["🔱 Godfather", "👨‍💼 CEO", "🚀 Moon Rocket", "🏆 Legend", "👑 Admin"]
        for role_name in cool_roles:
            role = discord.utils.find(lambda r: r.name == role_name, guild.roles)
            if role and role not in interaction.user.roles:
                await interaction.user.add_roles(role)
        
        embed = Embed(title="✅ Roles Setup Complete!", color=COLORS["green"])
        embed.add_field(
            name="🎖️ Your Roles",
            value="✅ 🔱 **Godfather**\n✅ 👨‍💼 **CEO**\n✅ 🚀 **Moon Rocket**\n✅ 🏆 **Legend**\n✅ 👑 **Admin**",
            inline=False
        )
        embed.add_field(
            name="📍 All Available Roles",
            value="\n".join(created_roles),
            inline=False
        )
        embed.add_field(
            name="🎯 How to Assign Roles to Members",
            value="Right-click member → **Roles** → Select the role\n\n💡 **Tip:** Roles appear as sections on the right side!",
            inline=False
        )
        await interaction.response.send_message(embed=embed, ephemeral=True)
        
    except Exception as e:
        embed = Embed(title="❌ Error", description=f"Failed to create roles: {str(e)}", color=COLORS["red"])
        await interaction.response.send_message(embed=embed, ephemeral=True)


if __name__ == "__main__":
    if not DISCORD_TOKEN:
        print("❌ DISCORD_TOKEN missing")
    else:
        client.run(DISCORD_TOKEN)
