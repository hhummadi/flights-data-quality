"""Data-quality rule engine for flights.csv (R01-R18).

Loads the dataset and computes a boolean failure mask per rule in ``F``
and the failure counts in ``fails``. The data path comes from the
``DQ_DATA`` environment variable, defaulting to ``data/raw/flights.csv``
and falling back to the bundled sample in ``data/sample/``.
"""
import os
from pathlib import Path
import pandas as pd, numpy as np

ROOT = Path(__file__).resolve().parent.parent
_default = ROOT / 'data' / 'raw' / 'flights.csv'
_sample = ROOT / 'data' / 'sample' / 'flights_sample.csv'
DATA_PATH = Path(os.environ.get('DQ_DATA') or (_default if _default.exists() else _sample))
REPORTS_DIR = Path(os.environ.get('DQ_REPORTS') or ROOT / 'reports')
df = pd.read_csv(DATA_PATH)
N=len(df)
def tomin(x): return (x//100)*60 + x%100
F={}
F['R01']=df.dep_time.isna()
F['R02']=df.dep_time.notna()&df.arr_time.isna()
F['R03']=df.arr_time.notna()&(df.arr_delay.isna()|df.air_time.isna())
F['R04']=df.tailnum.isna()
faa=df.tailnum.str.match(r'^N([1-9]\d{0,4}|[1-9]\d{0,3}[A-HJ-NP-Z]|[1-9]\d{0,2}[A-HJ-NP-Z]{2})$')
F['R05']=df.tailnum.notna()&~faa.fillna(False).astype(bool)
def badt(s): return s.notna()&((s%100>59)|(s>2359)|(s<0))
F['R06']=badt(df.dep_time)|badt(df.arr_time)|badt(df.sched_dep_time)|badt(df.sched_arr_time)
F['R07']=~df.carrier.isin(df.carrier.unique()) | ~df.origin.isin(['EWR','JFK','LGA']) | ~df.dest.str.match(r'^[A-Z]{3}$')
dt=pd.to_datetime(dict(year=df.year,month=df.month,day=df.day),errors='coerce')
F['R08']=dt.isna()|(df.year!=2013)
dd=tomin(df.dep_time)-tomin(df.sched_dep_time)
F['R09']=df.dep_delay.notna()&((dd-df.dep_delay).abs()==1440)
th=pd.to_datetime(df.time_hour,errors='coerce')
F['R10']=((df.hour*100+df.minute)!=df.sched_dep_time)|(th!=dt+pd.to_timedelta(df.hour,unit='h'))
F['R11']=df.carrier.map(df.groupby('carrier').name.nunique())>1
rd=df.groupby(['origin','dest']).distance.transform('nunique'); F['R12']=rd>1
tc=df.groupby('tailnum').carrier.transform('nunique'); F['R13']=tc.fillna(1)>1
el=(tomin(df.arr_time)-tomin(df.dep_time))%1440
F['R14']=df.air_time.notna()&(df.air_time>el)
sp=df.distance/(df.air_time/60); F['R15']=sp.notna()&((sp>600)|(sp<150))
x=df.dropna(subset=['tailnum']).copy(); x['sd']=th[x.index]+pd.to_timedelta(x.minute,unit='m')
x=x.sort_values(['tailnum','sd']); gap=x.groupby('tailnum').sd.diff().dt.total_seconds()/60; po=x.groupby('tailnum').origin.shift()
ovidx=x[(gap<60)&(po!=x.origin)].index
x['pidx']=x.groupby('tailnum').id.shift()
prev_of={i:int(x.loc[i,'pidx']) for i in ovidx}
F['R16']=df.index.isin(ovidx)
F['R17']=(df.dep_delay>600)|(df.arr_delay>600)
F['R18']=df.duplicated(['year','month','day','carrier','flight','origin'],keep=False)
df['speed_mph']=sp.round(1)
fails={k:int(v.sum()) for k,v in F.items()}
if __name__=='__main__':
    print(f'Dataset: {DATA_PATH} ({N:,} rows)')
    for k, v in fails.items():
        print(f'{k}: {v:>7,} failing rows ({(1 - v / N) * 100:6.2f}% pass)')
