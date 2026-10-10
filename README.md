# team-formation

Evidence-based student team formation for project courses. Students answer a short private survey (Microsoft Forms or Google Forms). Staff paste the class list and the survey export into an Excel workbook. A constraint programming model (Google OR-Tools CP-SAT) then assigns every student to a team and proves that the result is optimal.

Students do not choose their teams. Self selection among students who barely know each other produces homogeneous teams and leaves a group of students nobody chose (Oakley et al., 2004). Instead, the tool applies explicit, published criteria in a strict order of priority:

1. no team has exactly one woman or exactly one man (configurable),
2. every team shares enough free time to meet,
3. the team with the most skill gaps has as few as possible,
4. at most one student without a survey response per team,
5. as many teams as possible have a strong member for each skill,
6. a mix of disciplinary backgrounds,
7. members close to the skill threshold where no strong member exists.

It was first used to form the project teams of an Interaction Design module at UCL.

## Features

* **Every student is placed**, including those who did not answer the survey. Team sizes differ by at most one.
* **Strict priorities.** Each criterion is optimised in turn and then held at its best value (lexicographic optimisation), so skill coverage can never cost a team its shared free time.
* **Proven optimal.** CP-SAT proves each stage optimal, or the Report says it was stopped at the time limit.
* **Reproducible.** `--deterministic` gives the same teams for the same input, seed and OR-Tools version. The Report records the command, SHA-256 fingerprints of the code, and the library versions.
* **Careful linking.** Survey responses are linked to the class list by email, then by name in any word order (`Lin, Mei` links to Mei Lin). A name that fits two students is never guessed; it is reported with instructions.
* **Staff overrides.** Pin a student to a team, for example after an agreed reasonable adjustment, or keep students apart. Keep apart requests from students and staff are hard constraints.
* **Audit report.** The output lists the result of every stage, every criterion measured on the final teams, warnings, and every match staff should check.
* **Staff friendly.** Everything happens in one Excel workbook with a settings sheet and the survey questions to copy. On Windows, drag the workbook onto `allocate.bat`.

## Quick start

You need Python 3.10 or newer.

```bash
git clone https://github.com/M-Colley/team-formation.git
cd team-formation
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt       # Windows: .venv\Scripts\python.exe
.venv/bin/python allocate_teams.py examples/example_intake.xlsx --out-dir out --deterministic
```

The example is a fictional cohort of 59 students. On Windows you can instead double click `setup_windows.bat` once, then drag a workbook onto `allocate.bat`. On macOS and Linux, run `chmod +x *.sh && ./setup_mac_linux.sh` once, then `./allocate.sh path/to/workbook.xlsx`.

Output of the example run:

```text
59 students (55 with a response) into 12 teams
Solver: OPTIMAL in 23.1 s | lone members 0 | missing shared blocks 0 | worst skill gaps 2
RESPONSE NOT USED: Wren Withdrawn <wren.withdrawn@example.edu>: no roster student with this email or name. ...
UNMATCHED: Hana Example: 'Someone Unknown'
Wrote out/example_intake_teams.xlsx
      out/example_intake_teams.csv
```

## Use it for your course

| File | Purpose |
|---|---|
| `team_formation.xlsx` | Empty staff workbook. Start with its *Start here* sheet. Never share a filled-in copy with students. |
| `allocate_teams.py` | Forms the teams from a workbook and writes the results |
| `build_workbook.py` | Regenerates the empty workbook, for example after changing `layout.py` |
| `layout.py` | Cell positions and defaults shared by both scripts |
| `docs/survey_template_microsoft_forms.pdf` | The survey as students see it |
| `examples/example_intake.xlsx` | A filled-in fictional workbook (`tests/make_example.py` rebuilds it) |

1. **Settings.** Set the team size (default 4 to 5), the strong skill threshold, the number of shared free time blocks each team needs, the lists of skills, disciplines, time blocks and gender options, the mentors and the deadline.
2. **Survey.** Create the survey listed on the *Form questions* sheet. For Microsoft Forms you can start from the ready-made template: [open it and choose *Duplicate it*](https://forms.cloud.microsoft/Pages/ShareFormPage.aspx?id=_oivH5ipW0yTySEKEdmlwsMs234pDzRDm53PQ_RJizRUODZJRFg5U0FZVVpWVFVER0hBUjRVOUVMMS4u&sharetoken=zwZvhE7fBUXZazQiGD5q) ([PDF preview](docs/survey_template_microsoft_forms.pdf)). Restrict it to your organisation and record names (Microsoft Forms) or verified emails (Google Forms). The question titles are what the allocator searches for, so keep them identical to *Settings* column Q.
3. **Data.** After the deadline, paste the class list from your learning platform (students only: full name and email) into *Roster* from row 3. Paste the survey export into *Responses* at A1, or pass the export file with `--responses`.
4. **Overrides (optional).** Pin students to teams or keep students apart.
5. **Run.** `python allocate_teams.py team_formation.xlsx --deterministic`. The script writes `<name>_teams.xlsx` and `<name>_teams.csv` next to the workbook and never modifies the input.
6. **Check.** Read the *Report* sheet first: warnings, responses it could not use and why, links and keep apart matches to check, and names it could not match.
7. **Publish.** Create the groups in your learning platform from the CSV (Team, Mentor, Full name, Email) and announce the teams together with the criteria.
8. **Delete.** Remove the survey data and the output files once the teams are final.

**Google Forms.** Download the responses as CSV. Checkbox answers separated by commas are understood, and grid questions (`Rate your experience in each area [User research ...]`) are matched by their statements. Google Forms has no *Completion time* column; set that title in *Settings* column Q to `Timestamp`, or leave it, in which case the later row counts as the latest submission.

### Options

```bash
python allocate_teams.py FILE.xlsx [--responses EXPORT.xlsx|.csv] [--teams N] [--out-dir DIR]
                                   [--time-limit SEC] [--deterministic] [--seed S] [--priority LIST]
```

* `--responses` reads a survey export file directly instead of the *Responses* sheet.
* `--teams N` overrides the number of teams. The default is *Settings!B5*, or the fewest teams at the maximum size.
* `--time-limit SEC` is the solver limit for each criterion (default 30). Most criteria are proven optimal within seconds.
* `--priority` sets the order of the criteria, for example `--priority lone,time,cover`. A criterion left out is not optimised but is still measured and reported.
* `--deterministic` makes the run reproducible. In this mode `--time-limit` is a machine independent work budget, so the wall clock time is longer. Use it for the run you publish and keep the Report sheet. If a criterion stops at the wall clock safety limit (20 times the budget) before its work budget is used, the Report warns that the run may not be reproducible. OR-Tools does not promise identical results across operating systems or processor types.

## How the allocation works

**Who is placed.** The class list is authoritative. Every student on it is placed, including those without a survey response. If the list is empty, all respondents are placed.

**Linking responses.** Responses are linked by email, then by name in any word order, ignoring commas, hyphens and apostrophes. Microsoft Forms records the account username as the email, which often differs from the address in the class list, so there most links are by name; the Report lists every name linked in a different spelling or word order. A name that fits more than one student (two students called Yu Wang, or Wei Li and Li Wei) is never guessed. Such a response, and one that matches nobody, is listed in the Report with the reason, and its answers are not used until staff put the student's class list email in its Email cell. When a student submits twice, the latest submission counts.

**Hard constraints.** Every student is in exactly one team. Team sizes differ by at most one, so 59 students form 11 teams of 5 and 1 team of 4. Pins hold. Keep apart requests hold for every name that can be matched.

**Criteria in order of priority.** The model is solved one criterion at a time. Each criterion is optimised and then held at its best value while the next one is optimised.

| Order | Criterion | Optimises |
|---|---|---|
| 1 | `lone` | fewest (team, gender) pairs with exactly one member of a gender marked *Yes* in Settings |
| 2 | `time` | fewest shared free time blocks missing below the required number, summed over teams |
| 3 | `worst` | fewest skill gaps in the team with the most gaps (max min fairness) |
| 4 | `nonresp` | fewest students without a response beyond the first in a team |
| 5 | `cover` | most (team, skill) pairs with a strong member (rating at or above the threshold) |
| 6 | `div` | most distinct disciplines, summed over teams |
| 7 | `partial` | most (team, skill) pairs with a member at the threshold minus one or above |

**Gender.** The lone member rule is on for *Woman* and *Man* and off for non-binary and self described answers. Applying it to a category with one or two members would force those students into the same team and could disclose their gender to teammates. An answer that starts like exactly one option, such as `Non-binary` for *Non-binary, or I describe my gender another way*, counts as that option and is listed in the Report. Any other answer is reported as a warning. The question is optional, and students without a response count as free in every time block and contribute no skills.

**Keep apart requests.** Names are matched by email, by name in any word order, by a close spelling, or by a unique first name. Each line or `;` part is first read as one name, so `Lee, Sam` is Sam Lee. Only if that fails is it split at commas, `/`, `&`, `+` and `and`, so `Alex Gray, Sam Lee` is two people. `N/A` and similar answers are ignored. The Report lists every request with the students it was matched to and marks close spellings and first name matches for checking.

**Column matching.** The allocator finds each column by the titles in *Settings* column Q, which are also the survey question titles. It tries an exact match, then a title starting with the text, then one containing it. A criterion whose column is missing is switched off and reported.

## Privacy and ethics

* The workbook and the outputs contain personal data. Keep them in a staff only location and delete them once the teams are final. The `.gitignore` keeps all spreadsheets and CSV files except the empty template and the fictional example out of the repository.
* The survey asks for no nationality or first language. Cultural diversity raises creativity and lowers cohesion (Stahl et al., 2010), so it offers no clear allocation target, and discipline already separates most same background clusters.
* Self ratings are used only to balance teams, never for marking. Tell students so in the survey description, together with the criteria (Michaelsen & Sweet, 2008).

## Rationale

Instructor formed teams with explicit criteria work better than self selected teams for students who do not know each other (Oakley et al., 2004); studies that favour self selection used cohorts who already knew each other (Bacon et al., 1999). Women in female minority engineering teams participate less and report more anxiety than women in female majority teams (Dasgupta et al., 2015), and numerically isolated members in general are treated as tokens (Kanter, 1977). Lack of shared free time is a frequent practical cause of team breakdown (Oakley et al., 2004). Heterogeneous grouping benefits low achieving students without harming high achieving students (Lou et al., 1996). Criteria based assignment follows CATME Team-Maker (Layton et al., 2010).

After formation, each team can write a team charter with its mentor (Mathieu & Rapp, 2009), and peer assessment of contribution reduces free riding (Loughry et al., 2007). Check your institution's regulations before announcing peer adjusted marks; `mod_peerwork` (Moodle) and WebPA implement the adjustment.

## Tests

```bash
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pytest -q
```

The tests run on GitHub Actions for Windows, macOS and Linux. They check, among other things, that the fictional cohort of 59 forms 11 teams of 5 and 1 of 4 with no lone woman or man and enough shared free time; that each criterion outranks the next, also in a case where a weighted sum gave up a free block to close nine skill gaps; that non-respondents are spread; that Microsoft Forms `.xlsx` exports and Google Forms CSV exports are read in full; that names link in any word order and ambiguous names are never guessed; that keep apart texts such as `Lee, Sam` and `N/A` are read correctly; that Report numbers are measured on the final teams; and that two `--deterministic` runs give identical teams.

## Citation and licence

MIT licence, © 2026 Mark Colley. If you use the tool in teaching or research, please cite it using the metadata in [`CITATION.cff`](CITATION.cff) (GitHub shows a *Cite this repository* button).

## References

* Bacon, D. R., Stewart, K. A., & Silver, W. S. (1999). Lessons from the best and worst student team experiences: How a teacher can make the difference. *Journal of Management Education, 23*(5), 467–488.
* Dasgupta, N., Scircle, M. M., & Hunsinger, M. (2015). Female peers in small work groups enhance women's motivation, verbal participation, and career aspirations in engineering. *PNAS, 112*(16), 4988–4993.
* Kanter, R. M. (1977). Some effects of proportions on group life: Skewed sex ratios and responses to token women. *American Journal of Sociology, 82*(5), 965–990.
* Layton, R. A., Loughry, M. L., Ohland, M. W., & Ricco, G. D. (2010). Design and validation of a web-based system for assigning members to teams using instructor-specified criteria. *Advances in Engineering Education, 2*(1), 1–28.
* Lou, Y., Abrami, P. C., Spence, J. C., Poulsen, C., Chambers, B., & d'Apollonia, S. (1996). Within-class grouping: A meta-analysis. *Review of Educational Research, 66*(4), 423–458.
* Loughry, M. L., Ohland, M. W., & Moore, D. D. (2007). Development of a theory-based assessment of team member effectiveness. *Educational and Psychological Measurement, 67*(3), 505–524.
* Mathieu, J. E., & Rapp, T. L. (2009). Laying the foundation for successful team performance trajectories: The roles of team charters and performance strategies. *Journal of Applied Psychology, 94*(1), 90–103.
* Michaelsen, L. K., & Sweet, M. (2008). The essential elements of team-based learning. *New Directions for Teaching and Learning, 2008*(116), 7–27.
* Oakley, B., Felder, R. M., Brent, R., & Elhajj, I. (2004). Turning student groups into effective teams. *Journal of Student Centered Learning, 2*(1), 9–34.
* Perron, L., & Didier, F. CP-SAT. Google OR-Tools. https://developers.google.com/optimization
* Stahl, G. K., Maznevski, M. L., Voigt, A., & Jonsen, K. (2010). Unraveling the effects of cultural diversity in teams: A meta-analysis of research on multicultural work groups. *Journal of International Business Studies, 41*(4), 690–709.
