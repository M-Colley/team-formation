"""Cell positions and defaults shared by build_group_formation.py and allocate_teams.py.

Change a position here and both scripts follow. Rebuild the workbook afterwards.
"""

# Sheet names
START, FORM, SETTINGS, ROSTER, RESPONSES, OVERRIDES = (
    "Start here", "Form questions", "Settings", "Roster", "Responses", "Overrides")

# Settings: single values (column B)
MAX_SIZE, MIN_SIZE, N_TEAMS, THRESHOLD, REQ_BLOCKS, DEADLINE = 3, 4, 5, 6, 7, 8
ROSTER_COUNT, RESPONSE_COUNT, TEAMS_EFFECTIVE, SIZES_TEXT = 10, 11, 12, 13

# Settings: tables (header on TABLE_HEAD, values from TABLE_FIRST downwards)
TABLE_HEAD, TABLE_FIRST = 15, 16
TEAM_COL, MENTOR_COL, TEAM_SLOTS = 1, 2, 40          # A, B
SKILL_LABEL_COL, SKILL_STMT_COL, SKILL_SLOTS = 4, 5, 12  # D, E
DISC_COL, DISC_SLOTS = 7, 12                         # G
BLOCK_COL, BLOCK_SLOTS = 9, 21                       # I
GENDER_COL, GENDER_RULE_COL, GENDER_SLOTS = 11, 12, 8  # K, L
TA_COL, TA_SLOTS = 14, 10                            # N
KEY_LABEL_COL, KEY_TEXT_COL = 16, 17                 # P, Q

# Column keys: how the allocator finds each column in the Forms export (one row each, in order)
KEYS = [
    ("name", "Name", "Recorded automatically by Microsoft Forms"),
    ("email", "Email", "Recorded automatically by Microsoft Forms"),
    ("completed", "Completion time", "Recorded automatically; used to keep the latest submission"),
    ("discipline", "Which discipline is your previous degree in?", "Question title"),
    ("skills", "Rate your experience in each area", "Question title (statements come from column E)"),
    ("availability", "When are you usually free for team meetings outside class?", "Question title"),
    ("gender", "Gender (optional)", "Question title"),
    ("apart", "Is there anyone you should not be placed with? (optional, confidential)",
     "Question title"),
]

# Roster and Overrides: header row 1, grey example row 2, data from row 3
DATA_FIRST = 3
ROSTER_ROWS, OVERRIDE_ROWS = 500, 100
EXAMPLE_PREFIX = "example"

# Defaults written into a new workbook
SKILLS = [  # (short label, statement shown in the form)
    ("Research", "User research (interviews, surveys)"),
    ("UI design", "Visual / UI design (Figma)"),
    ("Sketching", "Sketching and storyboarding"),
    ("Prototyping", "Prototyping (paper and digital)"),
    ("Programming", "Programming (web, Unity, Python)"),
    ("Physical computing", "Physical computing (Arduino)"),
    ("Evaluation", "Evaluation and data analysis"),
    ("Video", "Video production and editing"),
    ("Writing", "Academic writing and presenting"),
]
DISCIPLINES = [
    "Psychology / cognitive science",
    "Computer science / engineering",
    "Design / architecture / arts",
    "Business / management",
    "Social sciences / humanities",
    "Natural sciences / medicine",
    "Other",
]
# Example blocks: no afternoons on days with scheduled classes. Edit them in Settings column I.
NO_AFTERNOON = ("Monday", "Tuesday", "Thursday", "Friday")
BLOCKS = [f"{day} {part}" for day in ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")
          for part in ("morning", "afternoon", "evening")
          if not (day in NO_AFTERNOON and part == "afternoon")] + ["Saturday", "Sunday"]
GENDERS = [  # (option, avoid a lone member of this gender in a team)
    ("Woman", True),
    ("Man", True),
    ("Non-binary, or I describe my gender another way", False),
    ("Prefer not to say", False),
]
TAS = ["Mentor 1", "Mentor 2", "Mentor 3", "Mentor 4", "Mentor 5"]  # placeholders, edit in Settings
RATING_OPTIONS = ["1 no experience", "2", "3 done it in a project", "4", "5 could teach it"]
