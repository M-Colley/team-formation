"""Draw the README header: the team composition of the fictional example cohort.

Runs the allocator on examples/example_intake.xlsx (deterministic, seed 0) and writes
docs/header_{light,dark}.{svg,png}. Needs matplotlib on top of requirements.txt:

    python -m pip install matplotlib
    python docs/make_header.py
"""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.patches import Circle, FancyBboxPatch  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import allocate_teams as at  # noqa: E402

W, H = 1280, 640
GROUPS = ["Psychology / cognitive science", "Computer science / engineering", "Design / architecture / arts"]
GROUP_LABELS = ["Psychology", "Computer science", "Design", "Other backgrounds"]
THEMES = {   # validated categorical slots 1-3 (all pairs), neutral fold, ink ramp for skills
    "light": dict(surface="#fcfcfb", ink="#0b0b0b", ink2="#52514e", muted="#898781", hair="#e1e0d9",
                  series=["#2a78d6", "#eb6834", "#1baf7a"], other="#898781",
                  strong="#6f6e69", partial="#a3a19a", good="#0ca30c"),
    "dark": dict(surface="#1a1a19", ink="#ffffff", ink2="#c3c2b7", muted="#898781", hair="#383835",
                 series=["#3987e5", "#d95926", "#199e70"], other="#898781",
                 strong="#b5b4ac", partial="#6b6a65", good="#0ca30c"),
}
plt.rcParams.update({"font.family": ["Segoe UI", "Arial", "DejaVu Sans"],
                     "svg.fonttype": "path", "svg.hashsalt": "team-formation"})


def example_teams():
    p = at.read_problem(ROOT / "examples" / "example_intake.xlsx", None, None)
    res = at.solve(p, at.DEFAULT_PRIORITY, 30, True, 0)
    return p, res


def group_of(s) -> int:
    if not s.responded:
        return 4
    return GROUPS.index(s.discipline) if s.discipline in GROUPS else 3


def check(ax, x, y, color):
    ax.add_line(Line2D([x, x + 5, x + 15], [y, y + 5, y - 6], color=color, lw=2.6,
                       solid_capstyle="round", solid_joinstyle="round"))


def render(p, res, mode: str, out: Path):
    c = THEMES[mode]
    fig = plt.figure(figsize=(W / 100, H / 100), dpi=100, facecolor=c["surface"])
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.set_aspect("equal")
    ax.axis("off")
    text = lambda x, y, s, size, color, **kw: ax.text(x, y, s, fontsize=size, color=color, va="center", **kw)

    cfg, T = p.cfg, range(p.n_teams)
    teams = [sorted((s for s in p.students if s.team == g), key=group_of) for g in T]
    rows = at.team_rows(p)
    shared = [r["n_shared"] for r in rows]
    proven = all(st == "OPTIMAL" for _, st, _, _ in res.stages)

    # left: title, claim, results of this run
    text(64, 138, "team-formation", 34, c["ink"], fontweight="bold")
    text(64, 190, "Evidence-based student project teams,", 15, c["ink2"])
    text(64, 216, "optimised with constraint programming.", 15, c["ink2"])
    claims = [f"{len(p.students)} students into {p.n_teams} teams",
              "No team with a lone woman or man" if res.terms["lone"] == 0 else
              f"{res.terms['lone']} lone members",
              f"Every team shares ≥ {min(shared)} free time blocks",
              "At most one non-respondent per team" if res.terms["nonresp"] == 0 else
              f"{res.terms['nonresp']} extra non-respondents",
              "Each criterion proven optimal" if proven else "Best teams found within the time limit"]
    for k, line in enumerate(claims):
        y = 290 + 40 * k
        check(ax, 66, y + 1, c["good"])
        text(94, y, line, 13, c["ink"])
    text(64, 600, "Fictional example cohort  ·  github.com/M-Colley/team-formation", 10.5, c["muted"])

    # right: one column per team
    x0, pitch, cw = 616, 50, 42
    xs = [x0 + pitch * g for g in T]
    text(x0, 74, "Teams", 11, c["ink2"], fontweight="bold")
    for g in T:
        text(xs[g] + cw / 2, 100, str(g + 1), 10, c["muted"], ha="center")
    text(604, 162, "Members", 10, c["ink2"], ha="right")
    for g, mem in enumerate(teams):
        for r, s in enumerate(mem):
            cx, cy = xs[g] + cw / 2, 124 + 19.5 * r
            k = group_of(s)
            if k == 4:
                ax.add_patch(Circle((cx, cy), 7, facecolor=c["surface"], edgecolor=c["muted"], lw=1.6))
            else:
                ax.add_patch(Circle((cx, cy), 8, facecolor=(c["series"] + [c["other"]])[k], edgecolor="none"))

    gy, gh, gp = 258, 13, 25                  # thin cells: the gaps and partials are the story
    text(x0, gy - 16, "Best self-rating in the team", 11, c["ink2"], fontweight="bold")
    for k, label in enumerate(cfg.skill_labels):
        y = gy + gp * k
        text(604, y + gh / 2, label, 10, c["ink2"], ha="right")
        for g, mem in enumerate(teams):
            best = max(s.skills[k] for s in mem)
            box = dict(boxstyle="round,pad=0,rounding_size=3")
            if best >= cfg.threshold:
                ax.add_patch(FancyBboxPatch((xs[g], y), cw, gh, facecolor=c["strong"], edgecolor="none", **box))
            elif best >= cfg.threshold - 1:
                ax.add_patch(FancyBboxPatch((xs[g], y), cw, gh, facecolor=c["partial"], edgecolor="none", **box))
            else:
                ax.add_patch(FancyBboxPatch((xs[g] + 0.75, y + 0.75), cw - 1.5, gh - 1.5, facecolor="none",
                                            edgecolor=c["muted"], lw=1.2, **box))
    sy = gy + gp * len(cfg.skill_labels) + 18
    text(604, sy, "Shared free blocks", 10, c["ink2"], ha="right")
    for g in T:
        text(xs[g] + cw / 2, sy, str(shared[g]), 10.5, c["ink"], ha="center")

    # legend
    ly, lx = 532, 480
    for k, label in enumerate(GROUP_LABELS + ["No survey response"]):
        if k == 4:
            ax.add_patch(Circle((lx + 7, ly), 5.5, facecolor=c["surface"], edgecolor=c["muted"], lw=1.4))
        else:
            ax.add_patch(Circle((lx + 7, ly), 6.5, facecolor=(c["series"] + [c["other"]])[k], edgecolor="none"))
        t = text(lx + 20, ly, label, 10, c["ink2"])
        lx += 20 + t.get_window_extent(fig.canvas.get_renderer()).width + 22
    ly, lx = 560, 480
    for kind, label in (("strong", f"Best member rated {cfg.threshold}–5"),
                        ("partial", f"Best rated {cfg.threshold - 1}"), ("gap", f"Best rated 1–{cfg.threshold - 2}")):
        box = dict(boxstyle="round,pad=0,rounding_size=3")
        if kind == "gap":
            ax.add_patch(FancyBboxPatch((lx + 0.6, ly - 5.9), 19, 11.8, facecolor="none", edgecolor=c["muted"],
                                        lw=1.1, **box))
        else:
            ax.add_patch(FancyBboxPatch((lx, ly - 6.5), 20, 13, facecolor=c[kind], edgecolor="none", **box))
        t = text(lx + 28, ly, label, 10, c["ink2"])
        lx += 28 + t.get_window_extent(fig.canvas.get_renderer()).width + 22

    fig.savefig(out.with_suffix(".svg"), facecolor=c["surface"])
    fig.savefig(out.with_suffix(".png"), dpi=200, facecolor=c["surface"])
    plt.close(fig)


if __name__ == "__main__":
    p, res = example_teams()
    for mode in ("light", "dark"):
        render(p, res, mode, ROOT / "docs" / f"header_{mode}")
    print("wrote docs/header_light.{svg,png} and docs/header_dark.{svg,png}")
