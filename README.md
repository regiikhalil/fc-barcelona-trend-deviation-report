## Install and run

Requires Python 3.10 or newer.

```bash
python -m pip install requirements.txt
python barcelona_charts.py --excel "/path/to/Team Stats Barcelona.xlsx"
```
CSV columns: date, team, opponent, competition, goals, xg, xga, shots, sot, possession_share, loss_low, loss_medium, loss_high, recoveries_low, recoveries_medium, recoveries_high, area_entries, offsides, goals_against, shots_against, sota, def_duels_attempted, def_duels_won, def_duel_success, final_third_completed, passes_per_possession, ppda.

The script uses earlier matches as its baseline; radar percentages equal match value / earlier-match mean × 100. Opponent-offsides and PSG zone metrics use earlier opponents as their baseline.

## Assets and implementation

pandas handles data, NumPy handles geometry, and Matplotlib draws and exports the graphs. The polygon projection adapts Matplotlib's radar example: https://matplotlib.org/stable/gallery/specialty_plots/radar_chart.html

Roboto and Roboto Condensed are included with their license. Club crest images are sourced from ESPN; each original asset URL is retained in the Python source.

Python implementation and visual iteration were developed with assistance from OpenAI Codex, with analytical direction and review by me.
