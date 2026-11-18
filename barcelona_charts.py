#!/usr/bin/env python3
"""Standalone Barcelona charts, using pandas, NumPy and Matplotlib.

Run: python barcelona_charts.py --excel "/path/to/Team Stats Barcelona.xlsx"
Or:  python barcelona_charts.py --data /path/to/match_stats.csv --out charts
Supply your own data; this minimal package contains no match data.
"""
from __future__ import annotations
import argparse
from dataclasses import dataclass
from functools import lru_cache
import hashlib
import json
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.lines import Line2D
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib.patches import FancyBboxPatch, Rectangle, RegularPolygon
from matplotlib.path import Path as MplPath
from matplotlib.projections import register_projection
from matplotlib.projections.polar import PolarAxes
from matplotlib.spines import Spine
from matplotlib.transforms import Affine2D
from PIL import Image

BASE=Path(__file__).resolve().parent
BLUE='#004D98';GOLD='#EDBB00';RED='#DB0030';BURGUNDY='#A50044'
NAVY='#14293A';INK='#273E50';MUTED='#536A79';GRID='#9AAAB7'
PALE_GOLD='#FFF4CC';BORDER='#CCD6DE'
FIGSIZE=(14,9.4)
METRICS={
 'offense': [('final_third_completed','Final-third passes\ncompleted'),('area_entries','Penalty area\nentries'),('loss_low','Possession losses\nlow zone'),('loss_medium','Possession losses\nmedium zone'),('loss_high','Possession losses\nhigh zone'),('passes_per_possession','Passes per\npossession'),('possession_share','Possession\nshare')],
 'defense': [('ppda','PPDA'),('recoveries_high','High-zone\nrecoveries'),('def_duels_attempted','Defensive duels\nattempted'),('def_duels_won','Defensive duels\nwon'),('def_duel_success','Defensive duel\nsuccess rate'),('opponent_offsides','Opponent offsides\ncommitted')],
 'psg': [('recoveries_low','Recoveries\nlow zone'),('recoveries_medium','Recoveries\nmedium zone'),('recoveries_high','Recoveries\nhigh zone'),('loss_low','Possession losses\nlow zone'),('loss_medium','Possession losses\nmedium zone'),('loss_high','Possession losses\nhigh zone')]
}

@dataclass
class Data:
    barcelona: pd.DataFrame
    opponents: pd.DataFrame
    prior: pd.DataFrame
    prior_opponents: pd.DataFrame
    match: pd.Series
    psg: pd.Series


def from_excel(path: Path) -> pd.DataFrame:
    """Read the original workbook directly through pandas' Excel reader."""
    raw=pd.read_excel(path,sheet_name=0,header=None)
    columns={'goals':6,'xg':7,'shots':8,'sot':9,'possession_share':14,'loss_low':16,'loss_medium':17,'loss_high':18,'recoveries_low':20,'recoveries_medium':21,'recoveries_high':22,'area_entries':52,'offsides':59,'goals_against':60,'shots_against':61,'sota':62,'def_duels_attempted':64,'def_duels_won':65,'def_duel_success':66,'final_third_completed':91,'passes_per_possession':104,'ppda':108}
    records=[]
    for i in raw.index[raw[4].eq('Barcelona')]:
        if raw.loc[i,2] not in ('Spain. La Liga','Europe. UEFA Champions League'):continue
        for row_idx,other_idx in [(i,i+1),(i+1,i)]:
            row=raw.loc[row_idx]
            values={key:float(row.iloc[col]) for key,col in columns.items()}
            records.append(dict(date=raw.loc[i,0],team=row.iloc[4],opponent=raw.loc[other_idx,4],competition='UCL' if 'Champions' in str(raw.loc[i,2]) else 'La Liga',xga=float(raw.loc[other_idx,7]),**values))
    return pd.DataFrame(records)


def load_data(path: Path,match_date: str,excel: bool=False) -> Data:
    df=from_excel(path) if excel else pd.read_csv(path)
    df['date']=pd.to_datetime(df['date'])
    df['competition']=df['competition'].replace({'UEFA Champions League':'UCL','Europe. UEFA Champions League':'UCL','Spain. La Liga':'La Liga'})
    match_day=pd.Timestamp(match_date)
    df=df[df['date'].le(match_day)].copy()
    eligible=df['competition'].isin(['La Liga','UCL','Spain. La Liga','Europe. UEFA Champions League'])
    df=df[eligible & (df['team'].eq('Barcelona')|df['opponent'].eq('Barcelona'))]
    for col in df.columns.difference(['date','team','opponent','competition','source_file']):
        df[col]=pd.to_numeric(df[col],errors='raise')
    barca=df[df['team'].eq('Barcelona')].sort_values('date').reset_index(drop=True)
    # Merge on both date and opponent identity; never rely on workbook row order.
    paired=barca[['date','opponent']].merge(df[df['opponent'].eq('Barcelona')],left_on=['date','opponent'],right_on=['date','team'],validate='one_to_one',suffixes=('_fixture',''))
    prior=barca[barca['date'].lt(match_day)].copy()
    prior_opp=paired[paired['date'].lt(match_day)].copy()
    match=barca[barca['date'].eq(match_day)]
    psg=paired[paired['date'].eq(match_day)]
    if len(match)!=1 or len(psg)!=1 or match.iloc[0]['opponent']!='PSG':raise ValueError('Exactly one Barcelona vs PSG fixture is required.')
    if prior.empty:raise ValueError('At least one earlier competitive match is required.')
    if len(prior)!=len(prior_opp):raise ValueError('Unpaired baseline fixture.')
    if barca.duplicated('date').any():raise ValueError('Duplicate Barcelona fixture date.')
    if not np.allclose(barca['xga'],paired['xg']):raise ValueError('Barcelona xGA differs from opponent xG.')
    if not (barca['sota']<=barca['shots_against']).all():raise ValueError('Shots on target exceed total shots.')
    return Data(barca,paired,prior,prior_opp,match.iloc[0],psg.iloc[0])


class Style:
    def __init__(self,assets: Path):
        self.assets=assets
        for path in (assets/'fonts').glob('*.ttf'):fm.fontManager.addfont(path)
        self.reg=fm.FontProperties(fname=assets/'fonts/Roboto-Regular.ttf')
        self.bold=fm.FontProperties(fname=assets/'fonts/Roboto-Bold.ttf')
        self.cond=fm.FontProperties(fname=assets/'fonts/RobotoCondensed-Bold.ttf')
        self.manifest={'Elche': {'provider': 'ESPN', 'provider_team_id': '3751', 'provider_name': 'Elche', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/3751.png', 'file': 'elche.png'}, 'Athletic Club': {'provider': 'ESPN', 'provider_team_id': '93', 'provider_name': 'Athletic Club', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/93.png', 'file': 'athleticclub.png'}, 'Rayo Vallecano': {'provider': 'ESPN', 'provider_team_id': '101', 'provider_name': 'Rayo Vallecano', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/101.png', 'file': 'rayovallecano.png'}, 'Valencia': {'provider': 'ESPN', 'provider_team_id': '94', 'provider_name': 'Valencia', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/94.png', 'file': 'valencia.png'}, 'Feyenoord': {'provider': 'ESPN', 'provider_team_id': '142', 'provider_name': 'Feyenoord Rotterdam', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/142.png', 'file': 'feyenoord.png'}, 'Levante': {'provider': 'ESPN', 'provider_team_id': '1538', 'provider_name': 'Levante', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/1538.png', 'file': 'levante.png'}, 'Racing Santander': {'provider': 'ESPN', 'provider_team_id': '87', 'provider_name': 'Racing Santander', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/87.png', 'file': 'racingsantander.png'}, 'Sevilla': {'provider': 'ESPN', 'provider_team_id': '243', 'provider_name': 'Sevilla', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/243.png', 'file': 'sevilla.png'}, 'Getafe': {'provider': 'ESPN', 'provider_team_id': '2922', 'provider_name': 'Getafe', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/2922.png', 'file': 'getafe.png'}, 'Galatasaray': {'provider': 'ESPN', 'provider_team_id': '432', 'provider_name': 'Galatasaray', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/432.png', 'file': 'galatasaray.png'}, 'Real Betis': {'provider': 'ESPN', 'provider_team_id': '244', 'provider_name': 'Real Betis', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/244.png', 'file': 'realbetis.png'}, 'PSG': {'provider': 'ESPN', 'provider_team_id': '160', 'provider_name': 'Paris Saint-Germain', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/160.png', 'file': 'psg.png'}, 'Barcelona': {'provider': 'ESPN', 'provider_team_id': '83', 'provider_name': 'Barcelona', 'url': 'https://a.espncdn.com/i/teamlogos/soccer/500/83.png', 'file': 'barcelona.png'}}
        plt.rcParams.update({'font.family':'Roboto','font.size':12,'text.color':INK,'axes.labelcolor':MUTED,'xtick.color':MUTED,'ytick.color':MUTED,'axes.unicode_minus':False,'pdf.fonttype':42,'svg.fonttype':'none','savefig.facecolor':'white'})

    @lru_cache(maxsize=None)
    def crest(self,team: str):
        return np.asarray(Image.open(self.assets/'club-crests'/self.manifest[team]['file']).convert('RGBA'))

    def add_crest(self,fig,team,xy,size=30,coords=None,offset=(0,0)):
        arr=self.crest(team)
        box=OffsetImage(arr,zoom=size/max(arr.shape[:2]),dpi_cor=True)
        annotation=AnnotationBbox(box,xy,xybox=offset,xycoords=coords or fig.transFigure,boxcoords='offset points',frameon=False,pad=0,annotation_clip=False)
        annotation.set_clip_on(False)
        fig.add_artist(annotation)
        return annotation

    def header(self,title,subtitle):
        fig=plt.figure(figsize=FIGSIZE,facecolor='white')
        fig.text(.045,.955,title,fontproperties=self.cond,fontsize=32,color=NAVY,va='top')
        fig.text(.045,.902,subtitle,fontproperties=self.cond,fontsize=14,color=MUTED,va='top')
        return fig

    def radar_legend(self,fig,opponent=False):
        frame=FancyBboxPatch((.75,.755),.215,.11,boxstyle='round,pad=.006,rounding_size=.006',edgecolor=BORDER,facecolor='white',linewidth=.9,transform=fig.transFigure)
        fig.add_artist(frame)
        for y,col,ls in [(.831,RED,'-'),(.784,BLUE,'--')]:
            fig.add_artist(Line2D([.763,.788],[y,y],color=col,ls=ls,lw=2.3,transform=fig.transFigure))
        self.add_crest(fig,'PSG',(.818,.831),24)
        fig.text(.851,.831,'vs',fontproperties=self.bold,fontsize=14,ha='center',va='center')
        self.add_crest(fig,'Barcelona',(.888,.831),24)
        fig.text(.803,.784,'Opponent average' if opponent else 'Season average',fontproperties=self.bold,fontsize=13,va='center')
        return frame


def line_chart(data: Data,s: Style,defense=False):
    actual,expected=('goals_against','xga') if defense else ('goals','xg')
    title='xG Against vs. Goals Against' if defense else 'Attacking xG vs. Goals'
    subtitle=('FC Barcelona Seasonal xG Against Trend' if defense else 'FC Barcelona Seasonal xG Trend')+' | La Liga and UEFA Champions League'
    fig=s.header(title,subtitle)
    ax=fig.add_axes([.085,.255,.865,.505])
    df=data.barcelona;x=np.arange(len(df));a=df[actual].to_numpy();e=df[expected].to_numpy()
    ax.axvspan(len(df)-1-.44,len(df)-1+.44,color='#FFED02',alpha=.12,zorder=0)
    ax.fill_between(x,a,e,where=a>=e,interpolate=True,color=BLUE,alpha=.075)
    ax.fill_between(x,a,e,where=a<e,interpolate=True,color=GOLD,alpha=.075)
    ax.plot(x,a,color=BLUE,lw=2.3,marker='o',ms=5,label='Goals Against' if defense else 'Actual goals scored',zorder=3)
    ax.plot(x,e,color=GOLD,lw=2.3,ls='--',marker='s',ms=4.8,label='xG Against (xGA)' if defense else 'Expected goals (xG)',zorder=3)
    maximum=float(np.ceil(max(a.max(),e.max()))) if defense else 8
    ax.set_ylim(0,maximum);ax.set_yticks(np.arange(0,maximum+1,1 if defense else 2))
    ax.set_xlim(-.5,len(df)-.5);ax.set_xticks(x,df['date'].dt.strftime('%m/%d'))
    ax.tick_params(axis='both',length=0);ax.tick_params(axis='x',pad=10)
    for label in ax.get_xticklabels():label.set_fontproperties(s.bold);label.set_fontsize(12);label.set_color(INK)
    ax.get_xticklabels()[-1].set_color(RED)
    for label in ax.get_yticklabels():label.set_fontproperties(s.bold);label.set_fontsize(12)
    ax.set_ylabel('Goals Against / xGA' if defense else 'Goals / xG',fontproperties=s.cond,fontsize=14,labelpad=14)
    ax.grid(axis='y',color=GRID,lw=1,alpha=.7);ax.set_axisbelow(True)
    for spine in ax.spines.values():spine.set_visible(False)
    legend=ax.legend(loc='lower right',bbox_to_anchor=(1,1.065),ncol=2,frameon=True,borderpad=.7,columnspacing=2,handlelength=2,prop=fm.FontProperties(fname=s.bold.get_file(),size=13))
    legend.get_frame().set_edgecolor(BORDER);legend.get_frame().set_linewidth(.9);legend.get_frame().set_alpha(1)
    ax.annotate(f'{e[-1]:.2f}',(x[-1],e[-1]),xytext=(0,13),textcoords='offset points',ha='center',fontproperties=s.bold,fontsize=14,color=INK)
    ax.annotate(f'{a[-1]:g}',(x[-1],a[-1]),xytext=(0,-18),textcoords='offset points',ha='center',fontproperties=s.bold,fontsize=14,color=BLUE)
    for i,row in df.iterrows():
        s.add_crest(fig,row['opponent'],(i,0),30,ax.get_xaxis_transform(),(0,-46))
        ax.annotate('UCL' if 'Champions' in row['competition'] or row['competition']=='UCL' else 'La Liga',(i,0),xytext=(0,-78),textcoords='offset points',ha='center',fontproperties=s.bold,fontsize=11,color=MUTED,annotation_clip=False)
    fig.text(.517,.075,'Chronological Fixures',ha='center',fontproperties=s.cond,fontsize=15,color=MUTED)
    return fig,{'kind':'line','y_min':0,'y_max':maximum,'fixtures':len(df)}


def scatter_chart(data: Data,s: Style):
    fig=s.header('Shots Against vs. Shots on Target Against','FC Barcelona Defensive Profile | La Liga and UEFA Champions League')
    ax=fig.add_axes([.085,.13,.865,.65])
    df=data.barcelona
    xmin,xmax=df['shots_against'].min(),df['shots_against'].max()
    ymin,ymax=df['sota'].min(),df['sota'].max()
    xlo,xhi=xmin-.55,xmax+.55;ylo,yhi=ymin-.6,ymax+.6
    mx,my=data.prior[['shots_against','sota']].mean()
    for x0,y0,x1,y1,color in [(xlo,ylo,mx,my,BLUE),(mx,ylo,xhi,my,GOLD),(xlo,my,mx,yhi,GOLD),(mx,my,xhi,yhi,BURGUNDY)]:
        ax.add_patch(Rectangle((x0,y0),x1-x0,y1-y0,color=color,alpha=.045,zorder=0))
    ax.set(xlim=(xlo,xhi),ylim=(ylo,yhi),xticks=np.arange(xmin,xmax+1,2),yticks=np.arange(ymin,ymax+1,2))
    ax.grid(color=GRID,alpha=.8,lw=.9);ax.set_axisbelow(True)
    ax.axvline(mx,color=BLUE,lw=1.5,ls=(0,(5,3)),zorder=2);ax.axhline(my,color=BLUE,lw=1.5,ls=(0,(5,3)),zorder=2)
    for _,row in df.iterrows():s.add_crest(fig,row['opponent'],(row['shots_against'],row['sota']),30,ax.transData)
    ax.set_xlabel('Shots Against',fontproperties=s.cond,fontsize=15,labelpad=16)
    ax.set_ylabel('Shots on Target Against',fontproperties=s.cond,fontsize=15,labelpad=16)
    for lab in ax.get_xticklabels()+ax.get_yticklabels():lab.set_fontproperties(s.bold);lab.set_fontsize(12)
    ax.tick_params(length=0,pad=9)
    for spine in ax.spines.values():spine.set_visible(False)
    return fig,{'kind':'scatter','x_data_range':[float(xmin),float(xmax)],'y_data_range':[float(ymin),float(ymax)],'quadrants':[float(mx),float(my)]}


def radar_projection(n):
    """Polygon-grid polar projection, adapted from Matplotlib's radar example.
    https://matplotlib.org/stable/gallery/specialty_plots/radar_chart.html
    """
    class RadarTransform(PolarAxes.PolarTransform):
        def transform_path_non_affine(self,path):
            if path._interpolation_steps>1:path=path.interpolated(n)
            return MplPath(self.transform(path.vertices),path.codes)
    class RadarAxes(PolarAxes):
        name=f'barcelona_radar_{n}'
        PolarTransform=RadarTransform
        def _gen_axes_patch(self):return RegularPolygon((.5,.5),n,radius=.5,edgecolor='none')
        def _gen_axes_spines(self):
            spine=Spine(self,'circle',MplPath.unit_regular_polygon(n))
            spine.set_transform(Affine2D().scale(.5).translate(.5,.5)+self.transAxes)
            return {'polar':spine}
    register_projection(RadarAxes)
    return RadarAxes.name


def radar_values(data: Data,kind: str) -> pd.DataFrame:
    metrics=METRICS[kind];keys=[key for key,_ in metrics if key!='opponent_offsides']
    current=data.psg if kind=='psg' else data.match
    prior=data.prior_opponents if kind=='psg' else data.prior
    result=pd.DataFrame({'metric':[k for k,_ in metrics],'label':[v for _,v in metrics]}).set_index('metric')
    result['match_value']=current.reindex(result.index)
    result['prior_average']=prior.reindex(columns=keys).mean().reindex(result.index)
    if kind=='defense':
        result.loc['opponent_offsides',['match_value','prior_average']]=[data.psg['offsides'],data.prior_opponents['offsides'].mean()]
    result[['match_value','prior_average']]=result[['match_value','prior_average']].astype(float)
    if (result['prior_average']<=0).any():raise ValueError('Radar baseline must be positive.')
    result['percent_of_average']=result['match_value'].div(result['prior_average']).mul(100)
    result['baseline_percent']=100.
    result['baseline_population']='Previous opponents' if kind=='psg' else 'Barcelona'
    if kind=='defense':result.loc['opponent_offsides','baseline_population']='Previous opponents'
    result['profile']=kind
    return result


def radar_chart(data: Data,s: Style,kind: str):
    values=radar_values(data,kind)
    maximum={'offense':150,'defense':200,'psg':250}[kind]
    if (values['percent_of_average']>maximum).any():raise ValueError(f'{kind}: a value exceeds the chosen scale.')
    title={'offense':'Offensive Profile','defense':'Defensive Profile','psg':'PSG Recoveries and Losses'}[kind]
    subtitle=('PSG vs. Barcelona | Compared with Barcelona\'s previous 11 opponents' if kind=='psg' else 'FC Barcelona | Season average compared to match vs PSG | La Liga and UEFA Champions League')
    fig=s.header(title,subtitle);legend_frame=s.radar_legend(fig,kind=='psg')
    n=len(values);theta=np.linspace(0,2*np.pi,n,endpoint=False)
    ax=fig.add_axes([.245,.115,.51,.66],projection=radar_projection(n))
    ax.set_theta_zero_location('N');ax.set_theta_direction(-1);ax.set_ylim(0,maximum)
    ax.set_xticks(theta);ax.set_xticklabels([])
    ax.set_yticks(np.arange(50,maximum+1,50));ax.set_yticklabels([])
    ax.grid(color='#B5C2CC',lw=.7);ax.spines['polar'].set_edgecolor('#B5C2CC');ax.spines['polar'].set_linewidth(.7)
    closed=np.r_[theta,theta[0]];baseline=np.full(n+1,100.)
    vals=values['percent_of_average'].to_numpy();closed_vals=np.r_[vals,vals[0]]
    ax.fill(closed,baseline,color=BLUE,alpha=.40,zorder=2);ax.plot(closed,baseline,color=BLUE,ls='--',lw=1.8,zorder=3)
    ax.fill(closed,closed_vals,color=RED,alpha=.45,zorder=4);ax.plot(closed,closed_vals,color=RED,lw=2.3,zorder=5)
    ax.scatter(theta,vals,c=RED,s=25,zorder=6)
    offsets={
      'offense':[(0,12,'center'),(12,10,'left'),(-12,12,'center'),(8,10,'left'),(-12,-20,'right'),(-12,7,'right'),(-12,10,'right')],
      'defense':[(0,12,'center'),(12,8,'left'),(10,-18,'left'),(0,-23,'center'),(-12,-18,'right'),(-12,9,'right')],
      'psg':[(-12,11,'right'),(12,9,'left'),(10,10,'left'),(0,-24,'center'),(-12,-20,'right'),(-12,10,'right')]
    }
    metric_labels=[];badges=[]
    for i,((_,row),a,pct,(dx,dy,ha)) in enumerate(zip(values.iterrows(),theta,vals,offsets[kind])):
        x,y=np.sin(a),np.cos(a)
        ha_axis='center' if abs(x)<.1 else 'left' if x>0 else 'right'
        rr=maximum*(1.17 if i==0 and '\n' in row['label'] else 1.13)
        metric_labels.append(ax.text(a,rr,row['label'],ha=ha_axis,va='center',fontproperties=s.cond,fontsize=14,color=INK,clip_on=False))
        badges.append(ax.annotate(f'{pct:.1f}%',(a,pct),xytext=(dx,dy),textcoords='offset points',ha=ha,va='baseline',fontproperties=s.bold,fontsize=14,color=NAVY,bbox=dict(boxstyle='round,pad=.27',facecolor=PALE_GOLD,edgecolor='#C9BA82',lw=.6),zorder=8,annotation_clip=False))
    radial_labels=[]
    for v in [100,maximum]:
        a=-.8 if v==100 and kind=='psg' else -.5 if v==100 and kind=='defense' else -.16 if v==100 else .075
        radial_labels.append(ax.text(a,v,f'{v}%',ha='right' if v==100 else 'left',va='center',fontproperties=s.bold,fontsize=10,color=BLUE if v==100 else MUTED,bbox=dict(facecolor='white',edgecolor='none',pad=1.2),zorder=7,clip_on=False))
    fig.canvas.draw();renderer=fig.canvas.get_renderer()
    boxes=[t.get_bbox_patch().get_window_extent(renderer) for t in badges]
    labels=[t.get_window_extent(renderer) for t in metric_labels]
    problems=[]
    for i,box in enumerate(boxes):
        for j,other in enumerate(boxes[i+1:],i+1):
            if box.overlaps(other):problems.append(f'badge {i} overlaps badge {j}')
        for j,other in enumerate(labels):
            if box.overlaps(other):problems.append(f'badge {i} overlaps metric label {j}')
    tick_boxes=[t.get_bbox_patch().get_window_extent(renderer) for t in radial_labels]
    for i,box in enumerate(boxes):
        for j,other in enumerate(tick_boxes):
            if box.overlaps(other):problems.append(f'badge {i} overlaps radial label {j}')
    for artist in [*metric_labels,*badges,*radial_labels]:
        box=artist.get_window_extent(renderer)
        if box.x0<0 or box.y0<0 or box.x1>fig.bbox.x1 or box.y1>fig.bbox.y1:
            problems.append(f'text outside figure: {artist.get_text()}')
    if problems:raise ValueError(f'{kind} layout: '+', '.join(problems))
    return fig,{'kind':'radar','profile':kind,'maximum':maximum,'axes':n,'layout_collisions':problems},values


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    inputs=parser.add_mutually_exclusive_group(required=True)
    inputs.add_argument('--data',type=Path,help='Match-statistics CSV')
    inputs.add_argument('--excel',type=Path,help='Original positional TeamStats workbook')
    parser.add_argument('--assets',type=Path,default=BASE/'assets')
    parser.add_argument('--out',type=Path,default=BASE/'charts')
    parser.add_argument('--match-date',default='2026-10-20')
    parser.add_argument('--dpi',type=int,default=300)
    args=parser.parse_args();args.out.mkdir(parents=True,exist_ok=True)
    data=load_data(args.excel or args.data,args.match_date,excel=args.excel is not None);style=Style(args.assets)
    specs=[('01-attacking-xg-vs-goals',lambda:line_chart(data,style)),('02-xga-vs-goals-against',lambda:line_chart(data,style,True)),('03-shots-against-scatter',lambda:scatter_chart(data,style)),('04-offensive-profile',lambda:radar_chart(data,style,'offense')),('05-defensive-profile',lambda:radar_chart(data,style,'defense')),('06-psg-recoveries-losses',lambda:radar_chart(data,style,'psg'))]
    audit=[];radar_tables=[]
    with PdfPages(args.out/'barcelona-charts.pdf',metadata={'Title':'Barcelona performance charts | PSG comparisons','Author':'','Creator':'pandas and Matplotlib'}) as pdf:
        for name,build in specs:
            result=build();fig,info=result[:2]
            if len(result)==3:radar_tables.append(result[2])
            fig.savefig(args.out/f'{name}.png',dpi=args.dpi)
            fig.savefig(args.out/f'{name}.svg')
            pdf.savefig(fig)
            audit.append({'chart':name,**info});plt.close(fig)
    pd.concat(radar_tables).reset_index().to_csv(args.out/'radar-values.csv',index=False)
    data.barcelona.to_csv(args.out/'barcelona-fixtures.csv',index=False)
    report={'backend':'Matplotlib','pandas_version':pd.__version__,'matplotlib_version':matplotlib.__version__,'numpy_version':np.__version__,'match_date':args.match_date,'baseline_matches':len(data.prior),'fixtures':len(data.barcelona),'prior_competitions':data.prior['competition'].value_counts().to_dict(),'source':str(args.excel or args.data),'source_sha256':hashlib.sha256((args.excel or args.data).read_bytes()).hexdigest(),'charts':audit}
    (args.out/'validation.json').write_text(json.dumps(report,indent=2))
    print(json.dumps({'pdf':str(args.out/'barcelona-charts.pdf'),'charts':len(specs),'baseline_matches':len(data.prior),'libraries':['pandas','numpy','matplotlib']}))

if __name__=='__main__':main()
