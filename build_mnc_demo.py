from pathlib import Path
import json, math, os, textwrap, zipfile, warnings
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler, MinMaxScaler
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.model_selection import train_test_split
from sklearn.metrics import roc_auc_score, classification_report
from sklearn.linear_model import LogisticRegression
import joblib

SCRIPT_DIR=Path(__file__).resolve().parent
if (SCRIPT_DIR/'10_README.md').exists() and (SCRIPT_DIR/'data').exists():
    ROOT=SCRIPT_DIR
elif (SCRIPT_DIR.parent/'10_README.md').exists() and (SCRIPT_DIR.parent/'data').exists():
    ROOT=SCRIPT_DIR.parent
else:
    ROOT=Path(os.environ.get('MNC_PROJECT_ROOT',SCRIPT_DIR/'MNC_Analytics_Demo'))
for d in ['data/raw','data/cleaned','charts','notebooks','app','outputs','reports']:
    (ROOT/d).mkdir(parents=True,exist_ok=True)
SEED=27102026
rng=np.random.default_rng(SEED)
END=pd.Timestamp('2025-12-31')
SYNTH='SYNTHETIC DEMONSTRATION — NOT REAL COMPANY DATA'

# -------------------------- synthetic source data --------------------------
def make_data():
    n=1800
    ids=np.arange(1,n+1)
    cities=np.array(['Mumbai','Delhi','Bengaluru','Hyderabad','Chennai','Pune','Kolkata','Ahmedabad'])
    city=rng.choice(cities,n,p=[.20,.17,.16,.12,.11,.10,.08,.06])
    seg=rng.choice(['Value','Core','Premium'],n,p=[.38,.44,.18])
    age=rng.normal(36,11,n).round().astype(float).clip(18,75)
    age[rng.choice(n,int(n*.025),replace=False)]=np.nan
    gender=rng.choice(['Female','Male','Non-binary','Prefer not to say'],n,p=[.48,.47,.02,.03])
    signup=pd.to_datetime(rng.integers(pd.Timestamp('2022-01-01').value//10**9, pd.Timestamp('2025-10-01').value//10**9,n),unit='s').normalize()
    cust=pd.DataFrame({'customer_id':[f'C{i:05d}' for i in ids],'age':age,'gender':gender,'city':city,'signup_date':signup,'segment':seg})
    cats=['Electronics','Home','Beauty','Fashion','Grocery']
    margins={'Electronics':.13,'Home':.34,'Beauty':.49,'Fashion':.39,'Grocery':.21}
    base={'Electronics':175,'Home':95,'Beauty':42,'Fashion':68,'Grocery':24}
    products=[]
    for i in range(1,41):
        c=cats[(i-1)%5]
        sub={'Electronics':['Audio','Accessories','Computing','Smart Home'],'Home':['Kitchen','Decor','Storage','Bedding'],'Beauty':['Skincare','Haircare','Makeup','Personal Care'],'Fashion':['Apparel','Footwear','Bags','Accessories'],'Grocery':['Pantry','Snacks','Beverages','Household']}[c][(i-1)//5%4]
        price=round(float(base[c]*rng.uniform(.55,2.2)),2)
        cost=round(price*(1-margins[c])*rng.uniform(.92,1.08),2)
        products.append([f'P{i:03d}',c,sub,cost,price])
    prod=pd.DataFrame(products,columns=['product_id','category','subcategory','cost_price','list_price'])
    cust_index=cust.set_index('customer_id')
    orders=[]
    oid=1
    seg_lambda={'Value':2.8,'Core':4.3,'Premium':6.1}
    category_weights={
      'Value':[.08,.17,.12,.18,.45], 'Core':[.18,.24,.16,.25,.17], 'Premium':[.38,.24,.14,.17,.07]}
    # Generate transactions by customer, always on/after signup.
    for ix,r in cust.iterrows():
        count=int(rng.poisson(seg_lambda[r.segment]))
        if count==0: continue
        start=max(pd.Timestamp(r.signup_date),pd.Timestamp('2024-01-01'))
        span=max(1,(END-start).days)
        for _ in range(count):
            dt=start+pd.Timedelta(days=int(rng.integers(0,span+1)))
            cat=rng.choice(cats,p=category_weights[r.segment])
            candidates=prod[prod.category.eq(cat)]
            p=candidates.iloc[int(rng.integers(0,len(candidates)))]
            qty=int(rng.choice([1,2,3,4],p=[.66,.22,.09,.03]))
            disc=round(float(rng.choice([0,.05,.10,.15,.20],p=[.43,.22,.20,.11,.04])),2)
            # Statuses deliberately include return/cancel/pending cases.
            status=rng.choice(['Completed','Returned','Cancelled','Pending'],p=[.855,.075,.04,.03])
            method=rng.choice(['Card','UPI','Wallet','COD'],p=[.32,.34,.13,.21])
            fail_p={'Card':.035,'UPI':.025,'Wallet':.055,'COD':.018}[method]
            pay_status='Failed' if rng.random()<fail_p else ('Refunded' if status=='Returned' else ('Pending' if status=='Pending' else 'Success'))
            gross=round(qty*float(p.list_price),2)
            discount=round(gross*disc,2)
            net=round(gross-discount,2)
            # City logistics quality creates a signal, but no causal interpretation is warranted.
            city_late={'Mumbai':.16,'Delhi':.21,'Bengaluru':.13,'Hyderabad':.18,'Chennai':.15,'Pune':.12,'Kolkata':.24,'Ahmedabad':.17}[r.city]
            late=bool(rng.random()<city_late)
            promised=dt+pd.Timedelta(days=int(rng.integers(3,7)))
            extra=int(rng.integers(1,5)) if late else int(rng.integers(-2,1))
            delivered=promised+pd.Timedelta(days=extra) if status in ['Completed','Returned'] else pd.NaT
            delivery_status=('Late' if late else 'On time') if status in ['Completed','Returned'] else ('In transit' if status=='Pending' else 'Not applicable')
            orders.append([f'O{oid:07d}',r.customer_id,dt,f'{p.product_id}',qty,float(p.list_price),disc,gross,discount,net,status,p.category,method,pay_status,promised,delivered,delivery_status,float(p.cost_price)])
            oid+=1
    orders=pd.DataFrame(orders,columns=['order_id','customer_id','order_date','product_id','qty','unit_price','discount_rate','gross_amount','discount_amount','order_value','order_status','category','payment_method','payment_status','promised_date','delivered_date','delivery_status','cost_price'])
    orders['revenue']=np.where((orders.order_status=='Completed')&(orders.payment_status=='Success'),orders.order_value,0.0)
    orders['profit']=np.where(orders.revenue>0,orders.revenue-orders.qty*orders.cost_price,0.0)
    orders['delivery_days']=(orders.delivered_date-orders.order_date).dt.days
    orders['late_flag']=(orders.delivery_status=='Late').astype(int)
    # Support table: complaint propensity rises with order volume and some segment/city differences.
    support=[]
    tid=1
    issue_types=['Delivery delay','Refund/payment','Product quality','Account/subscription','Other']
    issue_probs=[.30,.20,.22,.16,.12]
    for _,r in cust.iterrows():
        num=int(rng.poisson(.42 if r.segment!='Premium' else .56))
        for _ in range(num):
            date=pd.Timestamp(r.signup_date)+pd.Timedelta(days=int(rng.integers(0,max(1,(END-pd.Timestamp(r.signup_date)).days+1))))
            issue=rng.choice(issue_types,p=issue_probs)
            days=max(1,int(rng.gamma(2.2,1.4)+(2 if issue=='Refund/payment' else 0)))
            support.append([f'T{tid:06d}',r.customer_id,date,issue,days])
            tid+=1
    supp=pd.DataFrame(support,columns=['ticket_id','customer_id','ticket_date','issue_type','resolution_days'])
    # Snapshot outcome is generated from observed late-delivery, return and support signals plus baseline risk.
    per=orders.groupby('customer_id').agg(order_count=('order_id','count'),late_rate=('late_flag','mean'),return_rate=('order_status',lambda s:(s=='Returned').mean())).reindex(cust.customer_id).fillna(0)
    supagg=supp.groupby('customer_id').agg(ticket_count=('ticket_id','count'),avg_resolution_days=('resolution_days','mean')).reindex(cust.customer_id).fillna(0)
    cust=cust.merge(per,left_on='customer_id',right_index=True,how='left').merge(supagg,left_on='customer_id',right_index=True,how='left')
    logit=-2.65+1.35*cust.late_rate+1.2*cust.return_rate+0.11*cust.ticket_count+np.where(cust.segment.eq('Value'),.30,0)+np.where(cust.segment.eq('Premium'),-.25,0)
    p_churn=1/(1+np.exp(-logit))
    cust['churn_flag']=rng.binomial(1,p_churn)
    cust=cust.drop(columns=['order_count','late_rate','return_rate','ticket_count','avg_resolution_days'])
    payments=orders[['order_id','payment_method','payment_status','order_value']].rename(columns={'order_value':'payment_amount'}).copy()
    deliveries=orders[['order_id','promised_date','delivered_date','delivery_status']].copy()
    # Inject realistic quality defects into raw exports: duplicates, whitespace, and missing attributes.
    raw={'customers':cust.copy(),'orders':orders.copy(),'products':prod.copy(),'payments':payments.copy(),'delivery':deliveries.copy(),'support':supp.copy()}
    raw['customers'].loc[raw['customers'].sample(frac=.015,random_state=SEED).index,'city'] = raw['customers'].loc[raw['customers'].sample(frac=.015,random_state=SEED).index,'city'].astype(str)+' '
    raw['customers'].loc[raw['customers'].sample(frac=.008,random_state=SEED+1).index,'gender'] = None
    raw['orders'].loc[raw['orders'].sample(frac=.008,random_state=SEED+2).index,'category'] = raw['orders'].loc[raw['orders'].sample(frac=.008,random_state=SEED+2).index,'category'].astype(str)+' '
    raw['orders']=pd.concat([raw['orders'],raw['orders'].sample(frac=.006,random_state=SEED+3)],ignore_index=True)
    raw['support'].loc[raw['support'].sample(frac=.012,random_state=SEED+4).index,'issue_type']=raw['support'].loc[raw['support'].sample(frac=.012,random_state=SEED+4).index,'issue_type'].astype(str)+'  '
    for name,df in raw.items(): df.to_csv(ROOT/f'data/raw/{name}.csv',index=False)
    # Cleaning: trim categorical strings, deduplicate, parse dates, fill no customer demographics by explicit labels only where needed.
    clean={}
    for name,df in raw.items():
        z=df.copy()
        for c in z.select_dtypes(include='object').columns: z[c]=z[c].astype('string').str.strip()
        z=z.drop_duplicates().reset_index(drop=True)
        for c in z.columns:
            if c.endswith('_date') or c in ['order_date','promised_date','delivered_date','signup_date','ticket_date']:
                z[c]=pd.to_datetime(z[c],errors='coerce')
        clean[name]=z
    # Reconcile keyed fact table from cleaned orders and dimension tables.
    clean['orders']=clean['orders'].drop_duplicates('order_id').copy()
    clean['customers']['age']=pd.to_numeric(clean['customers']['age'],errors='coerce')
    clean['orders']['revenue']=np.where((clean['orders'].order_status=='Completed')&(clean['orders'].payment_status=='Success'),clean['orders'].order_value,0.0)
    clean['orders']['profit']=np.where(clean['orders'].revenue>0,clean['orders'].revenue-clean['orders'].qty*clean['orders'].cost_price,0.0)
    clean['orders']['delivery_days']=(clean['orders'].delivered_date-clean['orders'].order_date).dt.days
    clean['orders']['late_flag']=(clean['orders'].delivery_status=='Late').astype(int)
    return clean

D=make_data()
customers,orders,products,payments,delivery,support=[D[k] for k in ['customers','orders','products','payments','delivery','support']]
# Enriched analyst dataset, one row per order.
enriched=(orders.merge(customers[['customer_id','age','gender','city','signup_date','segment','churn_flag']],on='customer_id',how='left',validate='many_to_one')
          .merge(products[['product_id','subcategory']],on='product_id',how='left',validate='many_to_one'))
enriched['order_month']=enriched.order_date.dt.to_period('M').dt.to_timestamp()
enriched['day_of_week']=enriched.order_date.dt.day_name()
enriched['is_weekend']=enriched.order_date.dt.dayofweek.isin([5,6]).astype(int)
enriched['quarter']=enriched.order_date.dt.quarter
# Write cleaned sources and enriched modeling fact.
for name,df in D.items():
    df.to_csv(ROOT/f'data/cleaned/{name}.csv',index=False,date_format='%Y-%m-%d')
enriched.to_csv(ROOT/'data/cleaned/orders_enriched.csv',index=False,date_format='%Y-%m-%d')

# ----------------------------- analyses & charts ----------------------------
completed=enriched[enriched.order_status.eq('Completed') & enriched.payment_status.eq('Success')].copy()
orders_n=len(orders); revenue=float(orders.revenue.sum()); profit=float(orders.profit.sum()); aov=revenue/max(1,int((orders.revenue>0).sum()))
return_rate=float((orders.order_status=='Returned').mean())
settled=orders[orders.delivery_status.isin(['Late','On time'])]
on_time=float((settled.delivery_status=='On time').mean())
churn_rate=float(customers.churn_flag.mean())
late_churn=customers[['customer_id','churn_flag']].merge(enriched.groupby('customer_id').agg(any_late=('late_flag','max')),on='customer_id',how='left').fillna({'any_late':0})
late_churn['any_late']=late_churn['any_late'].astype(int)
ct=pd.crosstab(late_churn.any_late,late_churn.churn_flag)
late_group=float(late_churn.loc[late_churn.any_late==1,'churn_flag'].mean())
on_time_group=float(late_churn.loc[late_churn.any_late==0,'churn_flag'].mean())
churn_rr=(late_group/on_time_group) if on_time_group else np.nan
cat_perf=enriched.groupby('category',as_index=False).agg(revenue=('revenue','sum'),profit=('profit','sum'),orders=('order_id','count'))
seg_perf=enriched.groupby('segment',as_index=False).agg(revenue=('revenue','sum'),profit=('profit','sum'),orders=('order_id','count'))
city_perf=enriched.groupby('city',as_index=False).agg(revenue=('revenue','sum'),profit=('profit','sum'),orders=('order_id','count'),late_rate=('late_flag','mean'))
pay_perf=orders.assign(failure_or_refund=orders.payment_status.isin(['Failed','Refunded'])).groupby('payment_method',as_index=False).agg(payment_count=('order_id','count'),failed_or_refunded=('failure_or_refund','mean'),failure_rate=('payment_status',lambda s:(s=='Failed').mean()),refund_rate=('payment_status',lambda s:(s=='Refunded').mean()))
month_perf=orders.groupby(orders.order_date.dt.to_period('M').astype(str),as_index=False).agg(revenue=('revenue','sum'),orders=('order_id','count'))
issue_perf=support.groupby('issue_type',as_index=False).agg(tickets=('ticket_id','count'),avg_resolution_days=('resolution_days','mean'))
# RFM on customers at 2025-12-31 using valid completed orders only.
valid=orders[(orders.order_status=='Completed')&(orders.payment_status=='Success')]
rfm=valid.groupby('customer_id').agg(last_purchase=('order_date','max'),frequency=('order_id','nunique'),monetary=('revenue','sum')).reindex(customers.customer_id).fillna({'frequency':0,'monetary':0})
rfm['recency_days']=(END-rfm.last_purchase).dt.days
rfm['recency_days']=rfm.recency_days.fillna((END-pd.Timestamp('2024-01-01')).days)
rfm['spend_quartile']=pd.qcut(rfm.monetary.rank(method='first'),4,labels=['Q1 lowest','Q2','Q3','Q4 highest'])

# Charts
sns.set_theme(style='whitegrid',font_scale=.9)
def savefig(name):
    plt.tight_layout(); plt.savefig(ROOT/f'charts/{name}.png',dpi=160,bbox_inches='tight'); plt.close()
# 8 distinct analyses as a dashboard grid + individual PNGs
fig,axs=plt.subplots(2,4,figsize=(18,9))
sns.histplot(orders.loc[orders.revenue>0,'revenue'],bins=30,kde=True,ax=axs[0,0],color='#3b82f6'); axs[0,0].set(title='1. Recognized revenue per order',xlabel='Revenue ($)')
sns.boxplot(data=orders[orders.revenue>0],x='category',y='revenue',ax=axs[0,1],color='#93c5fd'); axs[0,1].tick_params(axis='x',rotation=35); axs[0,1].set(title='2. Revenue spread by category',ylabel='Revenue ($)')
sns.barplot(data=seg_perf,x='segment',y='revenue',ax=axs[0,2],color='#14b8a6'); axs[0,2].set(title='3. Revenue by segment',ylabel='Revenue ($)')
mp=month_perf.copy(); axs[0,3].plot(mp.order_date if 'order_date' in mp else range(len(mp)),mp.revenue,color='#7c3aed',marker='o',markersize=2); axs[0,3].set(title='4. Monthly recognized revenue',xlabel='Month',ylabel='Revenue ($)'); axs[0,3].tick_params(axis='x',labelrotation=65)
sns.scatterplot(data=customers.merge(rfm[['frequency','monetary']],left_on='customer_id',right_index=True),x='frequency',y='monetary',hue='segment',alpha=.5,ax=axs[1,0]); axs[1,0].set(title='5. Order frequency vs spend')
late_plot=late_churn.groupby('any_late').churn_flag.mean().reset_index(); sns.barplot(data=late_plot,x='any_late',y='churn_flag',ax=axs[1,1],color='#f97316'); axs[1,1].set(title='6. Churn rate by any late delivery',xlabel='Any late delivery (0/1)',ylabel='Synthetic churn rate')
issue_perf.sort_values('avg_resolution_days').plot.barh(x='issue_type',y='avg_resolution_days',ax=axs[1,2],legend=False,color='#64748b'); axs[1,2].set(title='7. Average support resolution',xlabel='Days',ylabel='Issue')
co=customers[['age','churn_flag']].copy(); co['age_band']=pd.cut(co.age,[17,25,35,45,55,100],labels=['18-25','26-35','36-45','46-55','56+']); heat=co.groupby('age_band',observed=False).churn_flag.mean().to_frame(); sns.heatmap(heat,annot=True,fmt='.1%',cmap='YlOrRd',ax=axs[1,3]); axs[1,3].set(title='8. Churn rate by age band')
fig.suptitle(SYNTH+' | EDA dashboard',fontsize=13,color='#b91c1c',fontweight='bold'); plt.tight_layout(rect=[0,0,1,.95]); fig.savefig(ROOT/'charts/08_chart_dashboard.png',dpi=160,bbox_inches='tight'); plt.close(fig)
# individual chart exports for selected figures
fig,ax=plt.subplots(figsize=(8,4)); sns.barplot(data=cat_perf.sort_values('revenue',ascending=False),x='category',y='revenue',ax=ax,color='#3b82f6'); ax.set(title='Synthetic recognized revenue by product category',ylabel='Revenue ($)'); ax.tick_params(axis='x',rotation=20); savefig('category_revenue')
fig,ax=plt.subplots(figsize=(8,4)); sns.barplot(data=late_plot,x='any_late',y='churn_flag',ax=ax,color='#f97316'); ax.set(title='Synthetic churn rate by whether customer had a late delivery',xlabel='Any late delivery',ylabel='Churn rate'); savefig('late_delivery_churn')
fig,ax=plt.subplots(figsize=(9,4)); ax.plot(month_perf.iloc[:,0],month_perf.revenue,marker='o',color='#7c3aed'); ax.set(title='Synthetic monthly recognized revenue',xlabel='Month',ylabel='Revenue ($)'); ax.tick_params(axis='x',rotation=70); savefig('monthly_revenue')

# ----------------------------- stats deliverable ----------------------------
manual=[
('A1.1 Mean, median and mode','Data: 120, 150, 150, 180, 200, 220, 250, 250, 300, 380.','Mean: formula = sum(x)/n → substitution = (120+150+150+180+200+220+250+250+300+380)/10 = 2,200/10 = 220. Answer: 220. Interpretation: average basket is $220 in this toy example.','Median: ordered middle values are 200 and 220; formula = (x5+x6)/2 → (200+220)/2 = 210. Answer: 210. Interpretation: half of these observations are at or below $210.','Mode: 150 and 250 each occur twice; all others occur once. Answer: bimodal, 150 and 250. Interpretation: two repeated spending points appear in this small sample.'),
('A1.2 Weighted mean','Formula = Σ(wᵢxᵢ)/Σwᵢ → (0.10×100 + 0.20×150 + 0.30×220 + 0.40×300)/(0.10+0.20+0.30+0.40) = (10+30+66+120)/1.00 = 226.','Answer: 226. Interpretation: the weighted average order value is $226; higher weights on larger baskets raise the result.'),
('A1.3 Median vs mean','Median is preferable when spending is skewed or contains a few unusually large orders. It describes a typical customer more robustly; mean remains useful for total-spend-per-customer planning.'),
('A2.4 Range, quartiles and IQR','Ordered data: 10,12,14,15,18,20,22,25,30,45. Range = max−min = 45−10 = 35. Using the median-of-halves convention: lower half 10,12,14,15,18 gives Q1=14; upper half 20,22,25,30,45 gives Q3=25. IQR=Q3−Q1=25−14=11.'),
('A2.5 1.5×IQR outlier rule','Lower fence = Q1−1.5×IQR =14−16.5=−2.5. Upper fence=Q3+1.5×IQR=25+16.5=41.5. Since 45>41.5, 45 is an outlier by this rule. Business meaning: investigate the unusually high value; do not automatically delete it.'),
('A2.6 Sample variance and standard deviation','Data 8,10,12,14,16; n=5; mean=60/5=12. Sample variance s²=Σ(x−x̄)²/(n−1)=[16+4+0+4+16]/4=40/4=10. Sample SD s=√10=3.1623. Answer: variance 10 units²; SD≈3.16 units.'),
('A2.7 High standard deviation','High SD means observations vary widely around the mean. For spending it can indicate heterogeneous customer behavior and make averages less representative; segmenting or using robust statistics may be more useful.'),
('A3.8 Z-score','z=(x−μ)/σ=(85−70)/10=15/10=1.5. Answer: 1.5 SD above the mean.'),
('A3.9 Compare unusual observations','For 660: z=(660−500)/80=160/80=2. For 350: z=(350−500)/80=−150/80=−1.875. Compare |z|: 2>1.875, so 660 is more unusual relative to this normal-scale reference.'),
('A3.10 Histogram and box plot','Histogram shows distribution shape, peaks, skew and gaps. Box plot summarizes median, quartiles and IQR and flags potential outliers, making group comparisons compact.'),
('A4.11 Payment probabilities','Online-paid=600/1000=0.60=60%; COD=250/1000=0.25=25%; wallet=150/1000=0.15=15%. They sum to 1.00.'),
('A4.12 Empirical return probability','P(return)=observed returns/observed orders=120/1000=0.12=12%. This estimates the rate under the sample period and process.'),
('A4.13 Three probability types','Classical: if each of four equally likely promo cards is equally likely to be drawn, P(one chosen card)=1/4. Empirical: return probability estimated as historical returns/orders. Subjective: a merchandiser’s forecast that a new product will sell through, based on experience and incomplete evidence.'),
('A5.14 Conditional late delivery among churned','P(Late|Churn)=180/300=0.60=60%.'),
('A5.15 Bayes: churn among late deliveries','P(Churn)=300/2000=0.15; P(Late|Churn)=180/300=0.60; P(Late|Not churn)=340/1700=0.20. Bayes: P(Churn|Late)=P(Late|Churn)P(Churn)/P(Late), where P(Late)=(180+340)/2000=0.26. Substitution=(0.60×0.15)/0.26=0.09/0.26=0.34615≈34.6%. Count check: 180/(180+340)=34.6%.'),
('A5.16 Conditional probability direction','P(A|B) conditions the denominator on B; P(B|A) conditions it on A. Their denominators/base rates differ, so they are generally unequal.'),
('A6.17 Binomial exactly two returns','For n=5, p=.20, k=2: P(X=2)=C(5,2)(.2)²(.8)³=10×.04×.512=.2048=20.48%, assuming independent orders with constant p.'),
('A6.18 Poisson exactly two tickets','λ=3 tickets/hour. P(X=2)=e^(−3)3²/2!=e^(−3)×4.5≈0.2240=22.40%, assuming a stable rate and independent arrivals.'),
('A6.19 Empirical rule','For an approximately normal distribution, about 68%, 95%, and 99.7% of observations lie within 1, 2, and 3 SDs of the mean, respectively. It is a rule of thumb, not universal.'),
('A7.20 Hypotheses','H0: μ_treatment−μ_control≤0 (new checkout does not increase AOV). H1: μ_treatment−μ_control>0 (it increases AOV). Define the population, exposure, and measurement window before testing.'),
('A7.21 Type I and II errors','Type I: conclude AOV increased when it did not (false positive), potentially rolling out an ineffective change. Type II: miss a real increase (false negative), potentially foregoing value.'),
('A7.22 p-value and tails','A p-value is the probability, assuming H0, of data at least as incompatible with H0 as observed. It is not the probability H0 is true. One-tailed tests a pre-specified direction; two-tailed tests either direction.'),
('A7.23 Test selection','Two-sample t-test: compare two independent group means (or paired t-test for paired observations), subject to assumptions/robust alternatives. ANOVA: compare means across 3+ groups. Chi-square: association between categorical variables or goodness-of-fit.'),
('A8.24 A/B groups','Control: eligible customers see existing checkout. Treatment: eligible customers see new checkout. Randomize at customer level and avoid cross-device contamination where feasible.'),
('A8.25 Primary and guardrail metrics','Primary: revenue per eligible randomized customer (or conversion if pre-registered). Guardrail: checkout error rate or refund rate. Choose one primary metric and define attribution window in advance.'),
('A8.26 Sample size and MDE','Sample size determines precision/power. Minimum detectable effect (MDE) is the smallest business-relevant effect the design aims to detect. Too-small samples can miss real effects; repeated peeking/multiple metrics inflate false positives unless controlled.'),
]
manual_md='# Manual Statistics & Probability\n\n'+SYNTH+' — the calculations below use the values stated in the assessment, not synthetic business records. Manual working is shown before software validation.\n\n'
for item in manual:
    manual_md+='## '+item[0]+'\n\n'+'\n\n'.join(item[1:])+'\n\n'
manual_md+='## Software validation after manual work\n\nThe central tendency, quartiles, sample variance, Z-scores and probability examples are independently re-computed in `06_Pandas_NumPy_EDA.ipynb`. Manual results above remain the primary answers.\n'
(ROOT/'reports/01_Manual_Statistics_Probability.md').write_text(manual_md)
# Create polished docx from markdown sections with python-docx.
def docx_from_sections(path,title,subtitle,sections):
    doc=Document(); sec=doc.sections[0]; sec.top_margin=Inches(.62); sec.bottom_margin=Inches(.62); sec.left_margin=Inches(.72); sec.right_margin=Inches(.72)
    styles=doc.styles; styles['Normal'].font.name='Aptos'; styles['Normal'].font.size=Pt(9)
    styles['Title'].font.name='Aptos Display'; styles['Title'].font.size=Pt(21); styles['Title'].font.color.rgb=RGBColor(31,78,121)
    styles['Heading 1'].font.name='Aptos Display'; styles['Heading 1'].font.size=Pt(14); styles['Heading 1'].font.color.rgb=RGBColor(31,78,121)
    styles['Heading 2'].font.name='Aptos'; styles['Heading 2'].font.size=Pt(11); styles['Heading 2'].font.color.rgb=RGBColor(31,78,121)
    p=doc.add_paragraph(); p.style='Title'; p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run(title)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run(subtitle); r.bold=True; r.font.color.rgb=RGBColor(185,28,28)
    for head,body in sections:
        doc.add_heading(head,level=1)
        for para in body.split('\n\n'):
            doc.add_paragraph(para)
    doc.save(path)

manual_sections=[]
for row in manual:
    manual_sections.append((row[0],'\n\n'.join(row[1:])))
manual_sections.append(('Validation and scope','Manual calculations above use assessment-provided values. The synthetic company dataset is separate and is used only for the portfolio analysis.'))
docx_from_sections(ROOT/'reports/01_Manual_Statistics_Probability.docx','Manual Statistics & Probability',SYNTH,manual_sections)

# ---------------------------- business findings -----------------------------
seg_top=seg_perf.sort_values('revenue',ascending=False).iloc[0]
cat_weak=cat_perf.assign(margin=cat_perf.profit/cat_perf.revenue.replace(0,np.nan)).sort_values('margin').iloc[0]
city_late=city_perf.sort_values('late_rate',ascending=False).iloc[0]
month_top=month_perf.sort_values('revenue',ascending=False).iloc[0]; month_low=month_perf.sort_values('revenue').iloc[0]
issue_top=issue_perf.sort_values('avg_resolution_days',ascending=False).iloc[0]
pay_top=pay_perf.sort_values('failed_or_refunded',ascending=False).iloc[0]
q4_customers=rfm[rfm.spend_quartile=='Q4 highest']
findings={
 'synthetic_label':SYNTH,'seed':SEED,'customers':len(customers),'orders':len(orders),'support_tickets':len(support),
 'recognized_revenue':revenue,'recognized_profit':profit,'recognized_orders':int((orders.revenue>0).sum()),'aov':aov,
 'return_rate':return_rate,'on_time_delivery_rate':on_time,'churn_rate':churn_rate,
 'late_delivery_customer_churn_rate':late_group,'no_late_delivery_customer_churn_rate':on_time_group,'relative_risk_late_vs_no_late':churn_rr,
 'top_revenue_segment':seg_top.segment,'segment_revenue':float(seg_top.revenue),'segment_profit':float(seg_top.profit),
 'weakest_category_margin':cat_weak.category,'weakest_category_margin_rate':float(cat_weak.margin),'weakest_category_revenue':float(cat_weak.revenue),'weakest_category_profit':float(cat_weak.profit),
 'top_spend_quartile_customers':int(len(q4_customers)),
 'highest_payment_composite_failure_refund_method':pay_top.payment_method,'highest_payment_failure_refund_rate':float(pay_top.failed_or_refunded),
 'highest_late_city':city_late.city,'highest_late_city_rate':float(city_late.late_rate),
 'strongest_month':str(month_top.iloc[0]),'strongest_month_revenue':float(month_top.revenue),'weakest_month':str(month_low.iloc[0]),'weakest_month_revenue':float(month_low.revenue),
 'slowest_support_issue':issue_top.issue_type,'slowest_support_avg_resolution_days':float(issue_top.avg_resolution_days),
}
(ROOT/'outputs/synthetic_findings.json').write_text(json.dumps(findings,indent=2,default=str))
cat_perf.to_csv(ROOT/'outputs/category_performance.csv',index=False); seg_perf.to_csv(ROOT/'outputs/segment_performance.csv',index=False); pay_perf.to_csv(ROOT/'outputs/payment_performance.csv',index=False); city_perf.to_csv(ROOT/'outputs/city_performance.csv',index=False); month_perf.to_csv(ROOT/'outputs/monthly_performance.csv',index=False); issue_perf.to_csv(ROOT/'outputs/support_performance.csv',index=False)

# ------------------------------- Excel files -------------------------------
def write_xlsx():
    # Cleaned workbook (source-aligned, cleaned table tabs + dictionary/quality).
    x1=ROOT/'02_Cleaned_Data.xlsx'
    with pd.ExcelWriter(x1,engine='xlsxwriter',datetime_format='yyyy-mm-dd',date_format='yyyy-mm-dd') as writer:
        for name,df in D.items(): df.to_excel(writer,sheet_name=name.title(),index=False)
        enriched.to_excel(writer,sheet_name='Orders_Enriched',index=False)
        q=[]
        for name,df in D.items(): q.append({'table':name,'rows':len(df),'columns':len(df.columns),'missing_cells':int(df.isna().sum().sum()),'duplicate_rows':int(df.duplicated().sum())})
        pd.DataFrame(q).to_excel(writer,sheet_name='Quality_Checks',index=False)
        dictionary=[]
        for name,df in D.items():
            for col in df.columns: dictionary.append({'table':name,'column':col,'dtype':str(df[col].dtype),'definition':'Synthetic field; see README for metric and generation assumptions.'})
        pd.DataFrame(dictionary).to_excel(writer,sheet_name='Data_Dictionary',index=False)
        wb=writer.book
        header=wb.add_format({'bold':True,'font_color':'white','bg_color':'#1F4E78','text_wrap':True,'valign':'vcenter','border':0})
        money=wb.add_format({'num_format':'$#,##0.00;($#,##0.00);-'}); pct=wb.add_format({'num_format':'0.0%'}); datefmt=wb.add_format({'num_format':'yyyy-mm-dd'})
        for s,ws in writer.sheets.items():
            ws.freeze_panes(1,0); ws.autofilter(0,0,ws.dim_rowmax,ws.dim_colmax); ws.set_row(0,28,header)
            ws.set_default_row(17)
            for ci,col in enumerate((D.get(s.lower(),enriched).columns if s not in ['Quality_Checks','Data_Dictionary'] else [])):
                ws.set_column(ci,ci,min(25,max(12,len(str(col))+2)))
            if s in ['Orders','Orders_Enriched']:
                for col in ['gross_amount','discount_amount','order_value','revenue','profit','unit_price','cost_price']:
                    if col in (D['orders'].columns if s=='Orders' else enriched.columns):
                        ci=(D['orders'] if s=='Orders' else enriched).columns.get_loc(col); ws.set_column(ci,ci,14,money)
            if s in ['Orders','Orders_Enriched']:
                for col in ['order_date','promised_date','delivered_date','signup_date']:
                    df=D['orders'] if s=='Orders' else enriched
                    if col in df.columns: ws.set_column(df.columns.get_loc(col),df.columns.get_loc(col),13,datefmt)
            if s in ['Quality_Checks','Data_Dictionary']:
                ws.set_column(0,0,22); ws.set_column(1,ws.dim_colmax,28)
    # Management dashboard workbook using formula-driven KPI controls and charted summary tables.
    x2=ROOT/'03_Excel_Management_Dashboard.xlsx'
    with pd.ExcelWriter(x2,engine='xlsxwriter',datetime_format='yyyy-mm-dd') as writer:
        enriched.to_excel(writer,sheet_name='Orders',index=False)
        month_perf.to_excel(writer,sheet_name='Monthly',index=False)
        cat_perf.to_excel(writer,sheet_name='Category',index=False)
        seg_perf.to_excel(writer,sheet_name='Segment',index=False)
        city_perf.to_excel(writer,sheet_name='City',index=False)
        pay_perf.to_excel(writer,sheet_name='Payments',index=False)
        issue_perf.to_excel(writer,sheet_name='Support',index=False)
        wb=writer.book; dash=wb.add_worksheet('Dashboard'); writer.sheets['Dashboard']=dash
        navy='#17365D'; teal='#0F766E'; blue='#2563EB'; pale='#EAF2F8'
        title=wb.add_format({'bold':True,'font_size':20,'font_color':'white','bg_color':navy,'align':'left','valign':'vcenter'})
        sub=wb.add_format({'bold':True,'font_color':'#B91C1C','bg_color':'#FEE2E2','align':'center','valign':'vcenter'})
        lab=wb.add_format({'bold':True,'font_color':'#334155','bg_color':pale,'border':1,'align':'center','valign':'vcenter'})
        val=wb.add_format({'bold':True,'font_size':17,'font_color':navy,'bg_color':'white','border':1,'num_format':'$#,##0.00'})
        val_pct=wb.add_format({'bold':True,'font_size':17,'font_color':navy,'bg_color':'white','border':1,'num_format':'0.0%'})
        controls=wb.add_format({'bg_color':'#FFF2CC','border':1,'align':'center'})
        dash.merge_range('A1:L1','MNC-READY ANALYTICS | MANAGEMENT DASHBOARD',title); dash.set_row(0,34)
        dash.merge_range('A2:L2',SYNTH,sub)
        dash.write('A4','Category filter',lab); dash.write('B4','All',controls)
        dash.write('C4','City filter',lab); dash.write('D4','All',controls)
        dash.write('E4','Segment filter',lab); dash.write('F4','All',controls)
        dash.write('G4','From date',lab); dash.write_datetime('H4',pd.Timestamp('2024-01-01').to_pydatetime(),controls)
        dash.write('I4','To date',lab); dash.write_datetime('J4',END.to_pydatetime(),controls)
        # hidden-ish lists source from lookup sheets; data validation direct list.
        dash.data_validation('B4',{'validate':'list','source':['All']+sorted(products.category.unique().tolist())})
        dash.data_validation('D4',{'validate':'list','source':['All']+sorted(customers.city.unique().tolist())})
        dash.data_validation('F4',{'validate':'list','source':['All','Value','Core','Premium']})
        dash.data_validation('H4',{'validate':'date','criteria':'between','minimum':datetime(2024,1,1),'maximum':datetime(2025,12,31)})
        dash.data_validation('J4',{'validate':'date','criteria':'between','minimum':datetime(2024,1,1),'maximum':datetime(2025,12,31)})
        # sheet columns
        cols={c:i for i,c in enumerate(enriched.columns)}
        def col_letter(idx):
            s=''; n=idx+1
            while n: n,rem=divmod(n-1,26); s=chr(65+rem)+s
            return s
        col={k:col_letter(v) for k,v in cols.items()}
        last=len(enriched)+1
        catcrit='IF($B$4="All","*",$B$4)'; citycrit='IF($D$4="All","*",$D$4)'; segcrit='IF($F$4="All","*",$F$4)'
        filters=f"Orders!${col['category']}$2:${col['category']}${last},{catcrit},Orders!${col['city']}$2:${col['city']}${last},{citycrit},Orders!${col['segment']}$2:${col['segment']}${last},{segcrit}"
        date_filter=f",Orders!${col['order_date']}$2:${col['order_date']}${last},\">=\"&$H$4,Orders!${col['order_date']}$2:${col['order_date']}${last},\"<=\"&$J$4"
        good_orders=f"Orders!${col['order_status']}$2:${col['order_status']}${last},\"Completed\",Orders!${col['payment_status']}$2:${col['payment_status']}${last},\"Success\",{filters}{date_filter}"
        completed_count=f"COUNTIFS(Orders!${col['order_status']}$2:${col['order_status']}${last},\"Completed\",Orders!${col['payment_status']}$2:${col['payment_status']}${last},\"Success\",{filters}{date_filter})"
        returned_count=f"COUNTIFS(Orders!${col['order_status']}$2:${col['order_status']}${last},\"Returned\",{filters}{date_filter})"
        all_count=f"COUNTIFS(Orders!${col['order_id']}$2:${col['order_id']}${last},\"<>\",{filters}{date_filter})"
        ontime_count=f"COUNTIFS(Orders!${col['delivery_status']}$2:${col['delivery_status']}${last},\"On time\",{filters}{date_filter})"
        late_count=f"COUNTIFS(Orders!${col['delivery_status']}$2:${col['delivery_status']}${last},\"Late\",{filters}{date_filter})"
        metric_specs=[('Revenue',f'=SUMIFS(Orders!${col["revenue"]}$2:${col["revenue"]}${last},{good_orders})',revenue),('Completed orders',f'={completed_count}',int((orders.revenue>0).sum())),('AOV','=IFERROR(B6/D6,0)',aov),('Return rate',f'=IFERROR({returned_count}/{all_count},0)',return_rate),('On-time delivery',f'=IFERROR({ontime_count}/({ontime_count}+{late_count}),0)',on_time)]
        positions=[('A6','B6'),('C6','D6'),('E6','F6'),('G6','H6'),('I6','J6')]
        for i,(label,formula,cached) in enumerate(metric_specs):
            lp,vp=positions[i]; dash.write(lp,label,lab); fmt=val_pct if label in ['Return rate','On-time delivery'] else val
            dash.write_formula(vp,formula,fmt,cached if label!='AOV' else (revenue/max(1,int((orders.revenue>0).sum()))))
        dash.write('A10','Dashboard view is formula-driven; change yellow controls to filter measures. Tables below are pivot-ready grouped summaries.',wb.add_format({'italic':True,'font_color':'#475569'}))
        # Charts from grouped summary sheets.
        c1=wb.add_chart({'type':'column'}); c1.add_series({'name':'Revenue','categories':['Category',1,0,len(cat_perf),0],'values':['Category',1,1,len(cat_perf),1],'fill':{'color':blue}}); c1.set_title({'name':'Recognized revenue by category (synthetic)'}); c1.set_y_axis({'name':'Revenue ($)'}); c1.set_legend({'none':True}); dash.insert_chart('A12',c1,{'x_scale':1.22,'y_scale':1.15})
        c2=wb.add_chart({'type':'bar'}); c2.add_series({'name':'Revenue','categories':['Segment',1,0,len(seg_perf),0],'values':['Segment',1,1,len(seg_perf),1],'fill':{'color':teal}}); c2.set_title({'name':'Revenue by customer segment'}); c2.set_x_axis({'name':'Revenue ($)'}); c2.set_legend({'none':True}); dash.insert_chart('G12',c2,{'x_scale':1.22,'y_scale':1.15})
        c3=wb.add_chart({'type':'line'}); c3.add_series({'name':'Revenue','categories':['Monthly',1,0,len(month_perf),0],'values':['Monthly',1,1,len(month_perf),1],'line':{'color':'#7C3AED'},'marker':{'type':'circle','size':4}}); c3.set_title({'name':'Monthly recognized revenue'}); c3.set_y_axis({'name':'Revenue ($)'}); c3.set_legend({'none':True}); dash.insert_chart('A31',c3,{'x_scale':1.55,'y_scale':1.15})
        c4=wb.add_chart({'type':'column'}); c4.add_series({'name':'Late rate','categories':['City',1,0,len(city_perf),0],'values':['City',1,4,len(city_perf),4],'fill':{'color':'#F97316'}}); c4.set_title({'name':'Late-delivery rate by city'}); c4.set_y_axis({'name':'Rate','num_format':'0%'}); c4.set_legend({'none':True}); dash.insert_chart('G31',c4,{'x_scale':1.22,'y_scale':1.15})
        notes=wb.add_worksheet('Build Notes'); writer.sheets['Build Notes']=notes
        notes.write_row(0,0,['Capability','Implementation / reproducible equivalent'],header)
        note_rows=[('Excel cleaning / formulas','Text cleanup: =TRIM(CLEAN(A2)); use Find & Replace for known text variants, then Data → Remove Duplicates on selected key(s). Enrichment: =XLOOKUP([@product_id],Products[product_id],Products[category],"Unmatched"); fallback =INDEX(Products[category],MATCH([@product_id],Products[product_id],0)). KPI examples: =SUMIFS(Orders[revenue],Orders[order_status],"Completed",Orders[payment_status],"Success"); =COUNTIFS(Orders[order_status],"Returned"); =IFERROR(ReturnedOrders/AllOrders,0).'),('Power Query / M','Applied-step recipe: Source → Promote Headers → Trim/Clean → Set Types → Remove Duplicates → Merge Orders with Customers/Products/Payments/Delivery → Add revenue/profit columns → Group by month/category → Load. M examples: Text.Trim(Text.Clean([city])); Table.Distinct(#"Trimmed Text"); Table.NestedJoin(Orders,{"customer_id"},Customers,{"customer_id"},"Customer",JoinKind.LeftOuter); Table.ExpandTableColumn(...); Table.Group(...). Workbook is generated via Python; M is guidance, not embedded query code.'),('PivotTables / slicers','Category/Segment/City/Monthly/Payments/Support sheets are pivot-ready group summaries and charts; native PivotTable caches/slicers are not authored by XlsxWriter. In desktop Excel, Insert → PivotTable from Orders, then add Category/City/Segment slicers and a Date timeline.'),('Power Pivot / DAX','Suggested star schema: FactOrders relates many-to-one to DimCustomer by customer_id, DimProduct by product_id, and DimDate by order_date. Revenue := CALCULATE(SUM(FactOrders[revenue]), FactOrders[order_status]="Completed", FactOrders[payment_status]="Success"); Profit := CALCULATE(SUM(FactOrders[profit]), FactOrders[order_status]="Completed"); AOV := DIVIDE([Revenue],[Completed Orders]); Revenue YTD := TOTALYTD([Revenue],DimDate[Date]); Return Rate := DIVIDE([Returned Orders],[All Orders]).'),('Controls','Dashboard category/city/segment dropdowns and date validation are active in this workbook. They provide formula-driven filter controls for headline KPIs.'),('Assumptions',SYNTH+'; revenue counts only Completed + Success; profit is recognized revenue less item cost, before overhead/returns handling; return rate is returned orders / all orders; on-time rate uses On time / (On time + Late).')]
        for r,(a,b) in enumerate(note_rows,1): notes.write(r,0,a); notes.write(r,1,b)
        notes.set_column(0,0,24); notes.set_column(1,1,110); notes.set_row(0,24,header)
        for wsname,ws in writer.sheets.items():
            if wsname!='Dashboard' and wsname!='Build Notes':
                ws.freeze_panes(1,0); ws.autofilter(0,0,ws.dim_rowmax,ws.dim_colmax); ws.set_row(0,24,header); ws.set_column(0,ws.dim_colmax,16)
        dash.set_column('A:L',15); dash.freeze_panes(5,0); dash.set_landscape(); dash.fit_to_pages(1,1); dash.set_margins(.25,.25,.4,.4)
    return x1,x2

# ------------------------------- SQL answers --------------------------------
sql=r'''/* 04_SQL_Assessment.sql
   PostgreSQL 15+ syntax. All tables are synthetic. Revenue definition: Completed and Success only.
   Assumed schema: customers(customer_id, age, gender, city, signup_date, segment, churn_flag)
   orders(order_id, customer_id, order_date, product_id, qty, unit_price, discount_rate,
          gross_amount, discount_amount, order_value, order_status, payment_method,
          payment_status, promised_date, delivered_date, delivery_status, cost_price, revenue, profit)
   products(product_id, category, subcategory, cost_price, list_price)
   payments(order_id, payment_method, payment_status, payment_amount)
   delivery(order_id, promised_date, delivered_date, delivery_status)
   support(ticket_id, customer_id, ticket_date, issue_type, resolution_days)
*/

-- 39. Completed orders, filtered and sorted.
SELECT order_id, customer_id, order_date, order_value
FROM orders WHERE order_status='Completed' AND payment_status='Success'
ORDER BY order_date DESC, order_id;

-- 40. Distinct cities and categories.
SELECT DISTINCT city FROM customers ORDER BY city;
SELECT DISTINCT category FROM products ORDER BY category;

-- 41. Basic aggregates on recognized revenue.
SELECT COUNT(*) AS order_rows, SUM(revenue) AS revenue, AVG(revenue) AS avg_order_revenue,
       MIN(revenue) AS min_order_revenue, MAX(revenue) AS max_order_revenue
FROM orders;

-- 42-43. Revenue by category; HAVING threshold shown as a parameter.
SELECT p.category, SUM(o.revenue) AS revenue
FROM orders o JOIN products p USING(product_id)
GROUP BY p.category
HAVING SUM(o.revenue) > :revenue_threshold
ORDER BY revenue DESC;

-- 44. ROUND / CEIL / FLOOR examples (Postgres uses CEIL; some engines use CEILING).
SELECT ROUND(AVG(revenue)::numeric,2) AS rounded_aov,
       CEIL(AVG(revenue)) AS aov_ceiling, FLOOR(AVG(revenue)) AS aov_floor FROM orders;

-- 45. Customers, orders and products: enriched order analysis.
SELECT c.customer_id,c.city,c.segment,o.order_id,o.order_date,p.category,p.subcategory,
       o.qty,o.order_value,o.revenue,o.profit
FROM customers c JOIN orders o USING(customer_id) JOIN products p USING(product_id);

-- 46. Customers who never ordered.
SELECT c.customer_id,c.city FROM customers c LEFT JOIN orders o USING(customer_id)
WHERE o.order_id IS NULL;

-- 47. Products never sold.
SELECT p.product_id,p.category,p.subcategory FROM products p LEFT JOIN orders o USING(product_id)
WHERE o.order_id IS NULL;

-- 48. Customer-level recognized revenue/profit, multiple joins.
SELECT c.customer_id,c.segment,c.city,SUM(o.revenue) AS revenue,SUM(o.profit) AS profit
FROM customers c JOIN orders o USING(customer_id) JOIN products p USING(product_id)
GROUP BY c.customer_id,c.segment,c.city ORDER BY revenue DESC;

-- 49. Self join: pair customers in same city, avoid duplicate/reversed pairs.
SELECT a.customer_id AS customer_a,b.customer_id AS customer_b,a.city
FROM customers a JOIN customers b ON a.city=b.city AND a.customer_id<b.customer_id
LIMIT 100;
-- Cross-join alternative for a small calendar/segment scaffold: SELECT s.segment,m.month FROM segments s CROSS JOIN months m;

-- 50. Customers with revenue above average customer revenue.
WITH customer_revenue AS (
 SELECT c.customer_id,SUM(o.revenue) AS revenue FROM customers c LEFT JOIN orders o USING(customer_id)
 GROUP BY c.customer_id
)
SELECT * FROM customer_revenue WHERE revenue>(SELECT AVG(revenue) FROM customer_revenue) ORDER BY revenue DESC;

-- 51. Monthly revenue CTE, with MoM change.
WITH monthly AS (
 SELECT DATE_TRUNC('month',order_date)::date AS month,SUM(revenue) AS revenue
 FROM orders GROUP BY 1
)
SELECT month,revenue,revenue-LAG(revenue) OVER(ORDER BY month) AS mom_change,
       revenue/NULLIF(LAG(revenue) OVER(ORDER BY month),0)-1 AS mom_pct
FROM monthly ORDER BY month;

-- 52. EXISTS / NOT EXISTS for customers with/without completed orders.
SELECT c.customer_id FROM customers c WHERE EXISTS
 (SELECT 1 FROM orders o WHERE o.customer_id=c.customer_id AND o.order_status='Completed');
SELECT c.customer_id FROM customers c WHERE NOT EXISTS
 (SELECT 1 FROM orders o WHERE o.customer_id=c.customer_id AND o.order_status='Completed');

-- 53. Recursive employee-manager hierarchy. Optional standalone schema/data, since employee data is not in the task source tables.
WITH RECURSIVE org AS (
 SELECT employee_id,manager_id,employee_name,1 AS depth,employee_name::text AS path
 FROM employees WHERE manager_id IS NULL
 UNION ALL
 SELECT e.employee_id,e.manager_id,e.employee_name,o.depth+1,o.path||' > '||e.employee_name
 FROM employees e JOIN org o ON e.manager_id=o.employee_id
)
SELECT * FROM org ORDER BY path;

-- 54. Latest order per customer.
SELECT * FROM (
 SELECT o.*,ROW_NUMBER() OVER(PARTITION BY customer_id ORDER BY order_date DESC,order_id DESC) AS rn
 FROM orders o
) q WHERE rn=1;

-- 55. Rank customers by recognized revenue within city.
WITH cr AS (SELECT c.customer_id,c.city,SUM(o.revenue) revenue FROM customers c LEFT JOIN orders o USING(customer_id) GROUP BY 1,2)
SELECT *,RANK() OVER(PARTITION BY city ORDER BY revenue DESC) AS revenue_rank,
       DENSE_RANK() OVER(PARTITION BY city ORDER BY revenue DESC) AS dense_rank
FROM cr;

-- 56. Spending quartiles.
WITH cr AS (SELECT customer_id,SUM(revenue) revenue FROM orders GROUP BY customer_id)
SELECT *,NTILE(4) OVER(ORDER BY revenue) AS spend_quartile FROM cr;

-- 57. LAG for MoM revenue, 58. LEAD for next transaction.
WITH m AS (SELECT DATE_TRUNC('month',order_date)::date month,SUM(revenue) revenue FROM orders GROUP BY 1)
SELECT month,revenue,LAG(revenue) OVER(ORDER BY month) previous_month FROM m;
SELECT customer_id,order_id,order_date,LEAD(order_date) OVER(PARTITION BY customer_id ORDER BY order_date) next_order_date
FROM orders;

-- 59-60. Running revenue and 3-month moving average.
WITH m AS (SELECT DATE_TRUNC('month',order_date)::date month,SUM(revenue) revenue FROM orders GROUP BY 1)
SELECT month,revenue,SUM(revenue) OVER(ORDER BY month ROWS UNBOUNDED PRECEDING) running_revenue,
       AVG(revenue) OVER(ORDER BY month ROWS BETWEEN 2 PRECEDING AND CURRENT ROW) moving_avg_3m FROM m;

-- 61-62. CASE classification and conditional aggregation.
SELECT customer_id,SUM(revenue) revenue,
 CASE WHEN SUM(revenue)>=5000 THEN 'High' WHEN SUM(revenue)>=1500 THEN 'Mid' ELSE 'Low' END value_band
FROM orders GROUP BY customer_id;
SELECT payment_method,COUNT(*) order_count,
 COUNT(*) FILTER(WHERE payment_status='Failed') failed,
 COUNT(*) FILTER(WHERE payment_status='Refunded') refunded,
 ROUND(100.0*COUNT(*) FILTER(WHERE payment_status IN ('Failed','Refunded'))/NULLIF(COUNT(*),0),2) issue_pct
FROM payments GROUP BY payment_method;

-- 63. String operations.
SELECT TRIM(city) AS clean_city,CONCAT(customer_id,'-',segment) AS customer_segment,
       SUBSTRING(customer_id FROM 2 FOR 5) AS id_suffix,REPLACE(segment,'Premium','VIP') AS display_segment
FROM customers;

-- 64. Date functions (PostgreSQL equivalents to DATEPART/DATEDIFF/DATEADD).
SELECT EXTRACT(YEAR FROM order_date) AS year,EXTRACT(QUARTER FROM order_date) AS quarter,
       (delivered_date::date-order_date::date) AS elapsed_days,
       (order_date + INTERVAL '30 days')::date AS date_plus_30 FROM orders;

-- 65. EXPLAIN / EXPLAIN ANALYZE (run in a safe test environment; ANALYZE executes query).
EXPLAIN (ANALYZE, BUFFERS) SELECT customer_id,SUM(revenue) FROM orders
WHERE order_date >= DATE '2025-01-01' GROUP BY customer_id;
-- Look for sequential scans on large facts, costly sorts/hash aggregates, row-estimate errors, spills and high buffer reads.

-- 66. Candidate B-tree / composite indexes; validate via workload and write-cost benchmark.
CREATE INDEX IF NOT EXISTS ix_orders_date_status ON orders(order_date,order_status);
CREATE INDEX IF NOT EXISTS ix_orders_customer_date ON orders(customer_id,order_date DESC);
CREATE INDEX IF NOT EXISTS ix_orders_product_status ON orders(product_id,order_status);
CREATE INDEX IF NOT EXISTS ix_support_customer_ticketdate ON support(customer_id,ticket_date);

-- 67. SARGability: keep indexed column bare in filter; use half-open date ranges.
SELECT * FROM orders WHERE order_date >= DATE '2025-01-01' AND order_date < DATE '2026-01-01';
-- Less SARGable: WHERE EXTRACT(YEAR FROM order_date)=2025; rewrite to range predicate above.

-- 68. EXISTS vs IN: EXISTS is a correlated existence test and handles NULL semantics predictably;
-- IN is concise for a set, but NULL in the subquery can make NOT IN evaluate UNKNOWN. Optimizers may transform either.
SELECT c.customer_id FROM customers c WHERE EXISTS (SELECT 1 FROM orders o WHERE o.customer_id=c.customer_id);

-- 69. Management sales view with grain made explicit.
CREATE OR REPLACE VIEW vw_management_sales AS
SELECT o.order_id,o.order_date,c.customer_id,c.city,c.segment,p.category,o.payment_method,
       o.revenue,o.profit,o.order_status,o.payment_status,o.delivery_status
FROM orders o JOIN customers c USING(customer_id) JOIN products p USING(product_id);

-- 70. RBAC / row-level security sketch (adapt roles and policy to platform/identity provider).
CREATE ROLE analyst_readonly;
GRANT SELECT ON vw_management_sales TO analyst_readonly;
-- Example RLS policy: CREATE POLICY city_scope ON customers USING (city=current_setting('app.city',true));
-- Restrict PII, grant least privilege, audit access, and enforce tenant filters in trusted session context.

-- 71. Platform fit: BigQuery serverless analytics with separation of storage/compute and on-demand or slot pricing;
-- Snowflake multi-cloud managed warehouse, elastic virtual warehouses and data sharing; Redshift AWS-integrated
-- warehouse with provisioned or serverless options. Choose based on ecosystem, workload, governance, skills, latency/cost.

-- 72. dbt-style plan: source declarations + freshness/unique/not_null/relationships tests; staging stg_orders,
-- stg_customers, stg_products; int_order_enriched; marts fct_orders/dim_customer/dim_product/dim_date;
-- tests for unique order_id, FK relationships, accepted statuses, nonnegative quantity/revenue; docs + lineage;
-- CI build changed models, snapshots for slowly changing dimensions, exposures for dashboard consumers.

-- 73. Storage choices: SQL for relational transactions and analytical joins/constraints; MongoDB for flexible
-- document-shaped application records with evolving schema; Redis for low-latency cache/session/queues/ephemeral state.
'''
(ROOT/'04_SQL_Assessment.sql').write_text(sql)

# ----------------------------- notebooks -----------------------------------
import nbformat as nbf

def write_nb(name,cells):
    nb=nbf.v4.new_notebook(); nb['cells']=[]; nb['metadata']={'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},'language_info':{'name':'python','version':'3.11'}}
    for typ,src in cells:
        if typ=='md': nb.cells.append(nbf.v4.new_markdown_cell(src))
        else: nb.cells.append(nbf.v4.new_code_cell(src))
    nbf.write(nb,ROOT/f'notebooks/{name}')

write_nb('05_Python_Foundations.ipynb',[
('md','# Python Foundations\n\n'+SYNTH+'\n\nRunnable from top to bottom from the project root or notebook folder; paths resolve from the notebook file.'),
('code',"""from pathlib import Path
import os, json, re
from datetime import datetime, date
import csv
BASE = Path.cwd()
if not (BASE/'data').exists(): BASE = BASE.parent
DATA = BASE/'data/cleaned'
print('Project root:', BASE.resolve())
"""),
('md','## Variables and core containers'),
('code',"""company = 'DemoCart'; order_count = 125; avg_basket = 84.5
items = ['orders','customers','products']
coordinates = (19.07,72.87)
unique_cities = {'Mumbai','Delhi','Pune'}
profile = {'segment':'Core','active':True,'orders':4}
print(company, order_count, avg_basket, items, coordinates, unique_cities, profile)
for item in items: print(item.upper())
price_doubles = [x*2 for x in [10,20,30]]
print(price_doubles)
"""),
('md','## Functions, *args, **kwargs, lambda, map and filter'),
('code',"""def line_total(price, qty=1, *discounts, **metadata):
    gross = price*qty
    return {'net': gross*(1-sum(discounts)), 'metadata':metadata}
print(line_total(20,2,.1,channel='web',sku='X1'))
amounts=[12,25,50,100]
print(list(map(lambda x: round(x*1.18,2), amounts)))
print(list(filter(lambda x: x>=25, amounts)))
"""),
('md','## os, json, datetime, regex and CSV/JSON I/O'),
('code',"""print('cwd:', os.getcwd(), 'files:', len(os.listdir(DATA)))
order_id='O0000123'; print(bool(re.fullmatch(r'O\\d{7}', order_id)))
print(datetime.now().date(), date(2025,12,31).isoformat())
record={'customer_id':'C00001','segment':'Core','spend':123.45}
json_path=BASE/'outputs/python_sample.json'; json_path.write_text(json.dumps(record,indent=2))
with json_path.open() as f: print(json.load(f))
with (BASE/'outputs/python_sample.csv').open('w',newline='') as f:
    w=csv.DictWriter(f,fieldnames=record.keys()); w.writeheader(); w.writerow(record)
with (BASE/'outputs/python_sample.csv').open() as f: print(list(csv.DictReader(f)))
"""),
('md','## CSV with Pandas and basic class / inheritance / encapsulation'),
('code',"""import pandas as pd
orders=pd.read_csv(DATA/'orders.csv',parse_dates=['order_date'])
print(orders.shape, orders[['order_id','customer_id','order_date']].head())
class Customer:
    def __init__(self, customer_id, segment): self.customer_id=customer_id; self._segment=segment
    @property
    def segment(self): return self._segment
class LoyaltyCustomer(Customer):
    def __init__(self, customer_id, segment, points): super().__init__(customer_id,segment); self.points=points
    def tier(self): return 'Gold' if self.points>=500 else 'Member'
c=LoyaltyCustomer('C00001','Core',620); print(c.customer_id,c.segment,c.tier())
"""),
('md','## Exceptions, custom exception and context manager'),
('code',"""class InvalidOrderError(Exception): pass
class ManagedTextFile:
    def __init__(self,path,mode='r',encoding='utf-8'): self.path=path; self.mode=mode; self.encoding=encoding; self.handle=None
    def __enter__(self): self.handle=open(self.path,self.mode,encoding=self.encoding); return self.handle
    def __exit__(self,exc_type,exc,tb): self.handle.close(); return False
try:
    qty=2
    if qty<=0: raise InvalidOrderError('Quantity must be positive')
    with ManagedTextFile(BASE/'outputs/python_sample.json') as f: print(json.load(f)['customer_id'])
except (InvalidOrderError, OSError, json.JSONDecodeError) as exc:
    print('Handled:',exc)
finally:
    print('Validation block completed')
""")])

write_nb('06_Pandas_NumPy_EDA.ipynb',[
('md','# Pandas, NumPy & Exploratory Data Analysis\n\n'+SYNTH+'\n\nThis workbook uses the local deterministic data under `../data/cleaned/`. CRISP-DM sequence: business question → data understanding → preparation → descriptive analysis → interpretation → communication. No causal claims are made.'),
('code',"""from pathlib import Path
import pandas as pd, numpy as np, matplotlib.pyplot as plt, seaborn as sns
BASE=Path.cwd(); BASE=BASE if (BASE/'data').exists() else BASE.parent
DATA=BASE/'data/cleaned'; OUT=BASE/'charts'; OUT.mkdir(exist_ok=True)
customers=pd.read_csv(DATA/'customers.csv',parse_dates=['signup_date'])
orders=pd.read_csv(DATA/'orders.csv',parse_dates=['order_date','promised_date','delivered_date'])
products=pd.read_csv(DATA/'products.csv')
payments=pd.read_csv(DATA/'payments.csv'); delivery=pd.read_csv(DATA/'delivery.csv')
support=pd.read_csv(DATA/'support.csv',parse_dates=['ticket_date'])
print('shape',orders.shape,'dtypes',orders.dtypes.to_dict())
print('nulls',orders.isna().sum().to_dict(),'duplicates',orders.duplicated().sum())
"""),
('md','## Import formats: CSV, Excel, JSON and SQL'),
('code',"""import sqlite3
excel_orders=pd.read_excel(BASE/'02_Cleaned_Data.xlsx',sheet_name='Orders')
json_metrics=pd.read_json(BASE/'outputs/synthetic_findings.json',typ='series')
conn=sqlite3.connect(':memory:')
orders.to_sql('orders',conn,index=False,if_exists='replace')
sql_sample=pd.read_sql_query('SELECT payment_method, COUNT(*) AS n FROM orders GROUP BY payment_method',conn)
print('Excel rows:',len(excel_orders),'JSON fields:',len(json_metrics),'SQL result rows:',len(sql_sample))
conn.close()
"""),
('md','## Clean, validate, merge and derive metrics'),
('code',"""for df in [customers,orders,products,payments,delivery,support]:
    for c in df.select_dtypes('object'): df[c]=df[c].str.strip()
orders=orders.drop_duplicates('order_id').copy()
orders['order_date']=pd.to_datetime(orders.order_date,errors='coerce')
orders['revenue']=np.where((orders.order_status=='Completed')&(orders.payment_status=='Success'),orders.order_value,0.0)
orders['profit']=np.where(orders.revenue>0,orders.revenue-orders.qty*orders.cost_price,0.0)
orders['delivery_days']=(orders.delivered_date-orders.order_date).dt.days
orders['late_flag']=(orders.delivery_status=='Late').astype(int)
orders['discount_amount']=orders.gross_amount*orders.discount_rate
fact=(orders.merge(customers,on='customer_id',how='left',validate='many_to_one')
      .merge(products[['product_id','subcategory']],on='product_id',how='left',validate='many_to_one'))
fact['month']=fact.order_date.dt.to_period('M').dt.to_timestamp()
print('unmatched customers',fact.age.isna().sum(),'duplicate order ids',fact.order_id.duplicated().sum())
"""),
('md','## Univariate and bivariate summaries'),
('code',"""display(orders[['order_value','revenue','profit','qty','discount_rate']].describe().T)
category=fact.groupby('category').agg(revenue=('revenue','sum'),profit=('profit','sum'),orders=('order_id','nunique')).sort_values('revenue',ascending=False)
segment=fact.groupby('segment').agg(revenue=('revenue','sum'),profit=('profit','sum'),orders=('order_id','nunique')).sort_values('revenue',ascending=False)
city=fact.groupby('city').agg(revenue=('revenue','sum'),late_rate=('late_flag','mean'),orders=('order_id','nunique'))
display(category,segment,city.sort_values('late_rate',ascending=False))
category_segment=pd.pivot_table(fact,index='category',columns='segment',values='revenue',aggfunc='sum',fill_value=0)
reconciled=pd.concat([orders.head(3),orders.tail(3)],ignore_index=True)
display(category_segment); print('concat sample rows',len(reconciled))
"""),
('md','## Eight-chart exploratory dashboard'),
('code',"""sns.set_theme(style='whitegrid')
fig,axs=plt.subplots(2,4,figsize=(18,9))
sns.histplot(orders.loc[orders.revenue>0,'revenue'],bins=30,kde=True,ax=axs[0,0]); axs[0,0].set_title('1. Revenue distribution')
sns.boxplot(data=fact[fact.revenue>0],x='category',y='revenue',ax=axs[0,1]); axs[0,1].tick_params(axis='x',rotation=30); axs[0,1].set_title('2. Revenue spread by category')
sns.barplot(data=segment.reset_index(),x='segment',y='revenue',ax=axs[0,2]); axs[0,2].set_title('3. Revenue by segment')
monthly=orders.groupby(orders.order_date.dt.to_period('M').astype(str)).revenue.sum(); axs[0,3].plot(monthly.index,monthly.values,marker='o'); axs[0,3].tick_params(axis='x',rotation=70); axs[0,3].set_title('4. Monthly revenue')
customer_spend=orders.groupby('customer_id').revenue.sum().rename('monetary').to_frame().join(orders.groupby('customer_id').order_id.nunique().rename('frequency'))
sns.scatterplot(data=customer_spend,x='frequency',y='monetary',ax=axs[1,0],alpha=.5); axs[1,0].set_title('5. Frequency vs spend')
late_customer=customers.set_index('customer_id')[['churn_flag']].join(orders.groupby('customer_id').late_flag.max().rename('any_late')).fillna({'any_late':0})
late_customer['any_late']=late_customer.any_late.astype(int)
sns.barplot(data=late_customer.groupby('any_late').churn_flag.mean().reset_index(),x='any_late',y='churn_flag',ax=axs[1,1]); axs[1,1].set_title('6. Churn by late-delivery exposure')
support.groupby('issue_type').resolution_days.mean().sort_values().plot.barh(ax=axs[1,2]); axs[1,2].set_title('7. Mean support resolution days')
tmp=customers[['age','churn_flag']].copy(); tmp['age_band']=pd.cut(tmp.age,[17,25,35,45,55,100]); sns.heatmap(tmp.groupby('age_band',observed=False).churn_flag.mean().to_frame(),annot=True,fmt='.1%',ax=axs[1,3]); axs[1,3].set_title('8. Churn by age band')
fig.suptitle('SYNTHETIC DEMONSTRATION — descriptive patterns, not causal evidence',color='darkred'); plt.tight_layout(); fig.savefig(OUT/'notebook_eda_dashboard.png',dpi=160,bbox_inches='tight'); plt.show()
"""),
('md','## Correlation and cross-tabulation: use with care'),
('code',"""numeric=fact[['qty','unit_price','discount_rate','revenue','profit','delivery_days','late_flag','churn_flag']].corr(numeric_only=True)
sns.heatmap(numeric,annot=True,fmt='.2f',cmap='vlag',center=0); plt.title('Synthetic numeric correlation matrix'); plt.show()
late_customer['churn_flag'].groupby(late_customer.any_late).agg(['mean','count'])
print(pd.crosstab(late_customer.any_late,late_customer.churn_flag,normalize='index'))
print('Association only: confounding and synthetic label generation preclude a causal interpretation.')
"""),
('md','## Manual-statistics validation (software only after manual answers)'),
('code',"""x=np.array([120,150,150,180,200,220,250,250,300,380])
print('mean',x.mean(),'median',np.median(x),'modes',pd.Series(x).mode().tolist())
y=np.array([10,12,14,15,18,20,22,25,30,45]); q1,q3=np.median(y[:5]),np.median(y[5:])
print('range',np.ptp(y),'Q1/Q3/IQR',q1,q3,q3-q1,'sample variance/std',np.var([8,10,12,14,16],ddof=1),np.std([8,10,12,14,16],ddof=1))
print('z', (85-70)/10,(660-500)/80,(350-500)/80)
print('manual probability checks',600/1000,120/1000,180/520,10*.2**2*.8**3,np.exp(-3)*3**2/2)
"""),
('md','## NumPy required operations'),
('code',"""a=np.arange(1,7,dtype=float); matrix=a.reshape(2,3)
print(a.shape,a.dtype,matrix.shape,matrix)
print('broadcast',matrix+np.array([10,20,30]))
print('sqrt/exp/log',np.sqrt(a),np.exp(a[:3]),np.log(a))
print('dot/matmul',np.dot([1,2],[3,4]),matrix@np.ones((3,1)))
print('boolean',a[a>3],'fancy',a[[0,2,5]],'where',np.where(a>=4,'high','low'))
""")])

write_nb('07_Feature_Engineering.ipynb',[
('md','# Feature Engineering & Preprocessing\n\n'+SYNTH+'\n\nCustomer-level snapshot features; churn_flag is the synthetic target. In production, define prediction timestamp and use only features observable before that timestamp to prevent target leakage.'),
('code',"""from pathlib import Path
import pandas as pd, numpy as np, joblib
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler, MinMaxScaler
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score, classification_report
BASE=Path.cwd(); BASE=BASE if (BASE/'data').exists() else BASE.parent
DATA=BASE/'data/cleaned'; OUT=BASE/'outputs'; OUT.mkdir(exist_ok=True)
c=pd.read_csv(DATA/'customers.csv',parse_dates=['signup_date'])
o=pd.read_csv(DATA/'orders.csv',parse_dates=['order_date'])
s=pd.read_csv(DATA/'support.csv',parse_dates=['ticket_date'])
"""),
('md','## Create at least five features: dates, spend, frequency, RFM, bins'),
('code',"""snapshot=pd.Timestamp('2025-12-31')
valid=o[(o.order_status=='Completed')&(o.payment_status=='Success')].copy()
rfm=valid.groupby('customer_id').agg(last_order=('order_date','max'),frequency=('order_id','nunique'),monetary=('revenue','sum'),monthly_spend=('revenue',lambda x:x.sum()/24)).reset_index()
rfm['recency_days']=(snapshot-rfm.last_order).dt.days
rfm['monthly_order_frequency']=rfm.frequency/24
late=o.groupby('customer_id').late_flag.mean().rename('late_delivery_rate')
returns=o.assign(is_return=(o.order_status=='Returned').astype(int)).groupby('customer_id').is_return.mean().rename('return_rate')
tickets=s.groupby('customer_id').agg(ticket_count=('ticket_id','count'),avg_resolution_days=('resolution_days','mean')).reset_index()
feat=(c.merge(rfm,on='customer_id',how='left').merge(late,on='customer_id',how='left').merge(returns,on='customer_id',how='left').merge(tickets,on='customer_id',how='left'))
for col in ['frequency','monetary','monthly_spend','recency_days','monthly_order_frequency','late_delivery_rate','return_rate','ticket_count','avg_resolution_days']:
    feat[col]=feat[col].fillna(0)
feat['signup_dow']=feat.signup_date.dt.day_name(); feat['signup_is_weekend']=feat.signup_date.dt.dayofweek.isin([5,6]).astype(int); feat['signup_quarter']=feat.signup_date.dt.quarter
feat['age_band']=pd.cut(feat.age,[17,25,35,45,55,100],labels=['18-25','26-35','36-45','46-55','56+'])
feat['spend_quartile']=pd.qcut(feat.monetary.rank(method='first'),4,labels=['Q1','Q2','Q3','Q4'])
feat['churn_flag']=feat.churn_flag.astype(int)
print(feat[['customer_id','recency_days','frequency','monetary','monthly_spend','monthly_order_frequency','late_delivery_rate','return_rate','ticket_count','age_band','signup_dow','signup_is_weekend','signup_quarter','spend_quartile']].head())
"""),
('md','## Train/test split, ColumnTransformer, Pipeline and persistence'),
('code',"""numeric=['age','recency_days','frequency','monetary','monthly_spend','monthly_order_frequency','late_delivery_rate','return_rate','ticket_count','avg_resolution_days','signup_is_weekend','signup_quarter']
categorical=['gender','city','segment','signup_dow','age_band','spend_quartile']
X=feat[numeric+categorical]; y=feat.churn_flag
pre=ColumnTransformer([('num',Pipeline([('imputer',SimpleImputer(strategy='median')),('scale',StandardScaler())]),numeric),('cat',Pipeline([('imputer',SimpleImputer(strategy='most_frequent')),('onehot',OneHotEncoder(handle_unknown='ignore'))]),categorical)])
model=Pipeline([('preprocess',pre),('classifier',LogisticRegression(max_iter=1000,class_weight='balanced',random_state=27102026))])
X_train,X_test,y_train,y_test=train_test_split(X,y,test_size=.25,random_state=27102026,stratify=y)
model.fit(X_train,y_train)
pred=model.predict_proba(X_test)[:,1]
print('test ROC-AUC (illustrative only):',round(roc_auc_score(y_test,pred),3))
print(classification_report(y_test,(pred>=.5).astype(int),zero_division=0))
joblib.dump(model,OUT/'churn_preprocessing_pipeline.joblib')
pre.fit(X_train); joblib.dump(pre,OUT/'preprocessor.joblib')
# Export full matrix transformed with training-fitted preprocessing (avoid refitting on holdout for evaluation).
X_trans=pre.transform(X); names=pre.get_feature_names_out()
mat=X_trans.toarray() if hasattr(X_trans,'toarray') else X_trans
ml=pd.DataFrame(mat,columns=names,index=feat.index); ml.insert(0,'customer_id',feat.customer_id); ml['churn_flag']=y
ml.to_csv(OUT/'ml_ready_customer_features.csv',index=False)
feat.to_csv(OUT/'customer_features.csv',index=False)
print('Saved pipeline and ML-ready export:',ml.shape)
""")])

# ------------------------------- Streamlit ----------------------------------
app=r'''"""Interactive dashboard for a deterministic synthetic e-commerce case study."""
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='MNC Analytics Demo',page_icon='📊',layout='wide')
BASE=Path(__file__).resolve().parents[1]
@st.cache_data
def load_data():
    o=pd.read_csv(BASE/'data/cleaned/orders_enriched.csv',parse_dates=['order_date'])
    s=pd.read_csv(BASE/'data/cleaned/support.csv')
    return o,s
o,s=load_data()
st.title('E-commerce Performance | Analytics Demo')
st.warning('SYNTHETIC DEMONSTRATION — all records and findings are simulated; do not interpret as real company performance.')
min_d=o.order_date.min().date(); max_d=o.order_date.max().date()
st.sidebar.header('Filters')
cats=st.sidebar.multiselect('Category',sorted(o.category.dropna().unique()),default=sorted(o.category.dropna().unique()))
cities=st.sidebar.multiselect('City',sorted(o.city.dropna().unique()),default=sorted(o.city.dropna().unique()))
segments=st.sidebar.multiselect('Segment',sorted(o.segment.dropna().unique()),default=sorted(o.segment.dropna().unique()))
dates=st.sidebar.date_input('Order date range',(min_d,max_d),min_value=min_d,max_value=max_d)
if len(dates)==2:
    start,end=pd.Timestamp(dates[0]),pd.Timestamp(dates[1])
    f=o[o.order_date.between(start,end+pd.Timedelta(days=1)-pd.Timedelta(microseconds=1))]
else: f=o.copy()
f=f[f.category.isin(cats)&f.city.isin(cities)&f.segment.isin(segments)]
recognized=f[(f.order_status=='Completed')&(f.payment_status=='Success')]
revenue=recognized.revenue.sum(); profit=recognized.profit.sum(); count=len(recognized); aov=revenue/count if count else 0
returns=(f.order_status=='Returned').mean() if len(f) else 0
ontime=f.delivery_status.isin(['On time','Late']); ontime_rate=(f.loc[ontime,'delivery_status'].eq('On time').mean() if ontime.any() else 0)
k=st.columns(5)
for col,label,value,fmt in zip(k,['Revenue','Profit','Completed orders','AOV','Return rate'],[revenue,profit,count,aov,returns],['${:,.0f}','${:,.0f}','{:,.0f}','${:,.2f}','{:.1%}']): col.metric(label,fmt.format(value))
st.caption(f'On-time delivery rate for observed delivered orders: {ontime_rate:.1%} | filtered records: {len(f):,}')
left,right=st.columns(2)
with left:
    agg=f.groupby('category',as_index=False).revenue.sum().sort_values('revenue',ascending=False)
    st.plotly_chart(px.bar(agg,x='category',y='revenue',title='Recognized revenue by category',labels={'revenue':'Revenue ($)'}),use_container_width=True)
with right:
    m=f.assign(month=f.order_date.dt.to_period('M').astype(str)).groupby('month',as_index=False).revenue.sum()
    st.plotly_chart(px.line(m,x='month',y='revenue',markers=True,title='Monthly recognized revenue',labels={'revenue':'Revenue ($)'}),use_container_width=True)
left,right=st.columns(2)
with left:
    seg=f.groupby('segment',as_index=False).agg(revenue=('revenue','sum'),profit=('profit','sum'))
    st.plotly_chart(px.bar(seg,x='segment',y=['revenue','profit'],barmode='group',title='Revenue and profit by segment'),use_container_width=True)
with right:
    city=f.groupby('city',as_index=False).late_flag.mean().sort_values('late_flag',ascending=False)
    st.plotly_chart(px.bar(city,x='city',y='late_flag',title='Late-delivery rate by city',labels={'late_flag':'Late rate'}),use_container_width=True)
st.subheader('Key demo insights')
st.write(f"Highest-revenue segment in the filtered view: **{seg.sort_values('revenue',ascending=False).iloc[0].segment if len(seg) else 'n/a'}**. This is descriptive and does not establish why performance differs.")
# Customer table, limit for responsive browser rendering.
st.subheader('Filtered order data')
st.dataframe(f[['order_id','order_date','customer_id','city','segment','category','payment_method','order_status','delivery_status','revenue','profit']].sort_values('order_date',ascending=False).head(1000),use_container_width=True)
st.download_button('Download filtered rows (CSV)',f.to_csv(index=False).encode('utf-8'),'filtered_orders_synthetic.csv','text/csv')
'''
(ROOT/'app/streamlit_app.py').write_text(app)
(ROOT/'app/requirements.txt').write_text('streamlit\npandas\nplotly\n')

# ---------------------------- executive report ------------------------------
exec_sections=[
('Scope and assumptions',f"This portfolio demonstrates an end-to-end e-commerce analytics workflow with a deterministic synthetic dataset (seed {SEED}), {len(customers):,} customers, {len(orders):,} orders and {len(support):,} support tickets covering 2024–2025. It is a demonstration only, not actual company data. Recognized revenue includes only Completed orders with Successful payment. Recognized profit subtracts item cost and excludes overhead. Return rate is returned orders divided by all orders. On-time delivery rate is On time divided by On time + Late. Churn is a synthetic customer snapshot outcome."),
('Executive snapshot',f"The demo produces recognized revenue of ${revenue:,.0f} and recognized profit of ${profit:,.0f}; {int((orders.revenue>0).sum()):,} completed/paid orders and AOV ${aov:,.2f}. Return rate is {return_rate:.1%}; on-time rate is {on_time:.1%}; synthetic churn prevalence is {churn_rate:.1%}. These values should be used to review analytical workflow, not as a forecast or benchmark."),
('Three findings',f"1. Segment performance: {seg_top.segment} is the highest-revenue segment (${seg_top.revenue:,.0f}) and contributes ${seg_top.profit:,.0f} recognized profit. 2. Category economics: {cat_weak.category} has the weakest observed contribution margin in this generated sample ({cat_weak.margin:.1%}; revenue ${cat_weak.revenue:,.0f}, profit ${cat_weak.profit:,.0f}); margin excludes overhead and is sensitive to simulated cost assumptions. 3. Delivery/churn association: customers with any recorded late delivery show synthetic churn rate {late_group:.1%}, compared with {on_time_group:.1%} among customers without a late delivery. The relative rate is {churn_rr:.2f}x; this is an association in generated data and is not causal evidence."),
('Operating signals',f"Payment method with the highest combined failed-or-refunded rate is {pay_top.payment_method} ({pay_top.failed_or_refunded:.1%}). City with the highest synthetic late-delivery rate is {city_late.city} ({city_late.late_rate:.1%}). Strongest/weakest monthly recognized revenue in the generated period is {month_top.iloc[0]} (${month_top.revenue:,.0f}) / {month_low.iloc[0]} (${month_low.revenue:,.0f}); seasonality in this small synthetic scenario is not a planning forecast. {issue_top.issue_type} has the longest mean support resolution ({issue_top.avg_resolution_days:.1f} days). The top-spending quartile contains {len(q4_customers):,} customers by rank-based quartile assignment."),
('Recommendations',"1. Prioritize a measured delivery-reliability pilot in the city with the highest late rate. Randomize eligible shipments or use phased rollout, track on-time rate and customer churn/retention, and monitor cost and complaint guardrails; do not infer causality from the observed association. 2. Review unit economics for the weakest-margin category: validate landed cost, discounts, return handling and vendor terms; test price/assortment interventions before broad rollout. 3. Investigate the payment method with the highest failed/refunded rate; segment failures vs refunds, audit gateway/error codes, and run a controlled payment UX or routing test with conversion and fraud/refund guardrails."),
('Modeling, caveats, next steps',"Useful candidate churn predictors include recency, order frequency, spend, change in order frequency, return/refund history, late-delivery exposure, support ticket volume and resolution time, payment failures, tenure, segment and city. Avoid leakage by fixing an observation window and predicting a later outcome window; evaluate time-based holdout, calibration, fairness, stability and business lift before operational use. The pipeline here is pedagogical and the synthetic label was generated using some correlated features, so its apparent predictive performance is not evidence of production model quality."),
('10-minute interview outline',"0–1 min: state the business problem and the synthetic-data caveat. 1–2 min: show source tables, cleaning checks and metric definitions. 2–4 min: present category/segment and month charts. 4–5 min: explain the manual sample SD or Bayes calculation. 5–6 min: walk through the customer-revenue CTE/window SQL. 6–7 min: show RFM features and the leakage-aware pipeline. 7–8 min: discuss the late-delivery/churn association without causal language. 8–9 min: deliver three recommendations. 9–10 min: state limitations and next-week priorities (real data validation, true Power Query/PivotTable model, experiment design, time-based model validation).")]
# report docx with report figures inserted after findings section
rep=Document(); sec=rep.sections[0]; sec.top_margin=Inches(.58); sec.bottom_margin=Inches(.58); sec.left_margin=Inches(.65); sec.right_margin=Inches(.65)
rep.styles['Normal'].font.name='Aptos'; rep.styles['Normal'].font.size=Pt(9)
rep.styles['Title'].font.name='Aptos Display'; rep.styles['Title'].font.size=Pt(22); rep.styles['Title'].font.color.rgb=RGBColor(31,78,121)
rep.styles['Heading 1'].font.name='Aptos Display'; rep.styles['Heading 1'].font.size=Pt(14); rep.styles['Heading 1'].font.color.rgb=RGBColor(31,78,121)
p=rep.add_paragraph(style='Title'); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run('Executive Analytics Report')
p=rep.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; rr=p.add_run(SYNTH); rr.bold=True; rr.font.color.rgb=RGBColor(185,28,28)
for i,(h,b) in enumerate(exec_sections):
    if h=='Operating signals': rep.add_page_break()
    rep.add_heading(h,level=1); rep.add_paragraph(b)
    if h=='Three findings':
        rep.add_picture(str(ROOT/'charts/category_revenue.png'),width=Inches(6.3))
        cp=rep.add_paragraph('Figure: Synthetic recognized revenue by category. See notebook for full eight-chart EDA dashboard.'); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER
rep.save(ROOT/'reports/09_Executive_Report.docx')

# interview talk track and README
readme=f'''# MNC-Ready Data Science & Data Analyst Project\n\n> **{SYNTH}.** This is a reproducible demonstration created only because synthetic data was explicitly requested. It is not a real company dataset and no real-world business conclusions should be drawn from it.\n\n## Deliverables\n\n- `01_Manual_Statistics_Probability.pdf` — manual-first calculations for every standalone statistics/probability question.\n- `02_Cleaned_Data.xlsx` — cleaned dimension/fact sheets, enriched order table, quality checks and data dictionary.\n- `03_Excel_Management_Dashboard.xlsx` — KPI formulas with dropdown/date controls, charts and pivot-ready group summaries. Includes Power Query steps/M snippets and DAX measures in Build Notes.\n- `04_SQL_Assessment.sql` — commented PostgreSQL solutions for the listed SQL topics.\n- `05_Python_Foundations.ipynb` — core Python, functions, file I/O, OOP, exceptions and context manager.\n- `06_Pandas_NumPy_EDA.ipynb` — profiling, cleaning, joins, feature derivation, manual-stat validation, required NumPy operations and eight-chart EDA.\n- `07_Feature_Engineering.ipynb` — RFM, monthly features, bins, encoding/scaling, ColumnTransformer, Pipeline, evaluation demo and joblib persistence.\n- `08_Streamlit_App/` — interactive dashboard app, data dependency requirements and this README's deployment instructions.\n- `09_Executive_Report.pdf` and `.docx` — summary, caveats, recommendations and 10-minute interview outline.\n- `10_README.md` — this file. Supporting CSVs, charts and outputs are included.\n\n## Synthetic data and assumptions\n\n- Fixed random seed: `{SEED}`. Period: 2024-01-01 to 2025-12-31. Tables: customers, orders, products, payments, delivery and support. Raw exports intentionally include small amounts of whitespace, missing demographics, and duplicate rows for cleaning practice.\n- `data/raw/` is deliberately imperfect. `data/cleaned/` is normalized, trimmed and deduplicated, with reproducible derived measures. No source system data is represented.\n- Recognized revenue = order value only when `order_status == Completed` and `payment_status == Success`. `profit` = recognized revenue − quantity × cost_price; it is a gross contribution proxy, not accounting profit. Returned orders are excluded from recognized revenue; returns remain in the denominator of overall return rate.\n- On-time delivery rate = On time / (On time + Late). Churn is a generated snapshot label at 2025-12-31. Synthetic churn probability was constructed with correlated late delivery/returns/support factors; the association analysis is descriptive, not causal.\n- Payment composite rate = Failed or Refunded / payment records. Failed and refunded rates are also reported separately.\n- RFM uses completed, successful orders as of 2025-12-31. Missing RFM activity is set to zero frequency/spend and long recency for a no-purchase customer.\n- Employee-manager hierarchy data was not in the suggested input schema; SQL includes a standalone recursive CTE template for an `employees` table.\n\n## Rebuild and run\n\nPython 3.10+ is recommended. From the project root:\n\n```bash\npip install pandas numpy matplotlib seaborn scikit-learn joblib openpyxl xlsxwriter nbformat nbclient ipykernel python-docx\npython build_mnc_demo.py\n````\n\nOpen notebooks in Jupyter and run top to bottom. Notebook data paths work from project root or `notebooks/`. `04_SQL_Assessment.sql` targets PostgreSQL; replace `:revenue_threshold` with a bound parameter in the client.\n\n## Streamlit Cloud deployment\n\nThe repository root should contain `app/streamlit_app.py`, `app/requirements.txt`, and the `data/cleaned/` directory. In Streamlit Community Cloud, choose `app/streamlit_app.py` as the app entry point and include project data files in the connected repository. From a local environment run `streamlit run app/streamlit_app.py`. The dashboard offers multi-select category/city/segment filters and date range controls, KPI cards, four interactive Plotly charts, a filtered dataframe and CSV download.\n\n## Excel feature notes\n\nThe dashboard workbook uses live Excel formulas (SUMIFS/COUNTIFS/IFERROR), list/date validation and native charts. XlsxWriter does not author native PivotTable caches/slicers or embed executable Power Query M queries. For transparency, the workbook includes pivot-ready grouped tables, a Power Query applied-step recipe/M snippets and DAX measure examples. In desktop Excel, create a native PivotTable from the `Orders` sheet table and add slicers/timeline if required. The formulas are pre-calculated/cached for the default all-data view; Excel recalculates after filter control changes.\n\n## Key demo findings\n\n- Recognized revenue: **${revenue:,.0f}**; recognized gross contribution proxy: **${profit:,.0f}**; AOV: **${aov:,.2f}**.\n- Highest revenue segment: **{seg_top.segment}**; revenue **${seg_top.revenue:,.0f}**, profit **${seg_top.profit:,.0f}**.\n- Lowest simulated category contribution margin: **{cat_weak.category}**, **{cat_weak.margin:.1%}**.\n- Synthetic churn: customers with any late delivery **{late_group:.1%}** vs **{on_time_group:.1%}** without; not causal.\n- Highest combined failed/refunded payment method: **{pay_top.payment_method}**, **{pay_top.failed_or_refunded:.1%}**.\n- Slowest-resolving support issue: **{issue_top.issue_type}**, mean **{issue_top.avg_resolution_days:.1f} days**.\n\n## QA and limitations\n\nData joins validate many-to-one dimensions; raw/clean table row counts and nulls are included in `02_Cleaned_Data.xlsx`. Statistical calculations are manually displayed in the PDF and recalculated in the notebook afterward. Business outputs are deterministic with the fixed seed. The logistic regression is only a learning exercise; its target construction induces feature correlations and does not establish generalization. Use real data, approved definitions, time-based validation, privacy review and experimentation before any operational decision.\n'''
(ROOT/'10_README.md').write_text(readme)
(ROOT/'08_Streamlit_App').mkdir(exist_ok=True)
# Put app code + requirements in named folder too, along with a copy of data it uses.
(ROOT/'08_Streamlit_App/streamlit_app.py').write_text(app)
(ROOT/'08_Streamlit_App/requirements.txt').write_text('streamlit\npandas\nplotly\n')
(ROOT/'08_Streamlit_App/README.md').write_text('Run from project root: `streamlit run 08_Streamlit_App/streamlit_app.py`. Data is in `data/cleaned/orders_enriched.csv` and `data/cleaned/support.csv`. For Streamlit Cloud, deploy this folder with the project `data/cleaned/` directory and requirements. '+SYNTH+'.\n')
# Primary standalone entry point at root-compatible app/ and requested folder.
(ROOT/'10_minute_interview_talk_track.md').write_text('# 10-minute interview talk track\n\n'+SYNTH+'\n\n'+'\n\n'.join(f'**{h}.** {b}' for h,b in exec_sections[-1:])+'\n')

# Build Excel after all tables/analysis
xclean,xdash=write_xlsx()

# PDF conversion for reports
os.makedirs(ROOT/'reports/pdf_tmp',exist_ok=True)
# PDF utility converted through LibreOffice headless in isolated profile.
import subprocess
for docx in [ROOT/'reports/01_Manual_Statistics_Probability.docx',ROOT/'reports/09_Executive_Report.docx']:
    subprocess.run(['soffice','-env:UserInstallation=file:///tmp/lo_mnc_demo_profile','--headless','--convert-to','pdf','--outdir',str(ROOT),str(docx)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
# Rename generated PDFs to requested top-level names.
for stem in ['01_Manual_Statistics_Probability','09_Executive_Report']:
    p=ROOT/f'{stem}.pdf'
    if not p.exists():
        alt=ROOT/f'{stem.replace("09_","09_")}.pdf'
        if alt.exists(): p=alt
# Copy canonical outputs to root using expected filenames (already named by input stem).
# Copy chart overview and model artefacts remain in charts/outputs.

# Validate notebooks by executing cells individually in a kernel via nbclient.
from nbclient import NotebookClient
for nbpath in (ROOT/'notebooks').glob('*.ipynb'):
    nb=nbf.read(nbpath,as_version=4)
    try:
        NotebookClient(nb,timeout=180,kernel_name='python3',resources={'metadata':{'path':str(ROOT)}}).execute()
        nbf.write(nb,nbpath)
        print('EXECUTED',nbpath.name)
    except Exception as e:
        print('NOTEBOOK_EXECUTION_WARNING',nbpath.name,repr(e))

import shutil
for notebook_name in ['05_Python_Foundations.ipynb','06_Pandas_NumPy_EDA.ipynb','07_Feature_Engineering.ipynb']:
    shutil.copy2(ROOT/'notebooks'/notebook_name,ROOT/notebook_name)
if Path(__file__).resolve() != (ROOT/'build_mnc_demo.py').resolve():
    shutil.copy2(Path(__file__).resolve(),ROOT/'build_mnc_demo.py')

# Include generated artifacts manifest.
manifest=[]
for f in sorted(ROOT.rglob('*')):
    if f.is_file() and f.name!='MANIFEST.txt': manifest.append(f'{f.relative_to(ROOT)}\t{f.stat().st_size} bytes')
(ROOT/'MANIFEST.txt').write_text('\n'.join(manifest)+'\n')
# Zip all project deliverables
zpath=ROOT.parent/'MNC_Analytics_Demo.zip'
with zipfile.ZipFile(zpath,'w',zipfile.ZIP_DEFLATED) as z:
    for f in ROOT.rglob('*'):
        if f.is_file(): z.write(f,Path(ROOT.name)/f.relative_to(ROOT))
print(json.dumps({'root':str(ROOT),'zip':str(zpath),'files':len(manifest),'customers':len(customers),'orders':len(orders),'revenue':revenue,'profit':profit,'churn':churn_rate,'late_churn':late_group,'no_late_churn':on_time_group},indent=2))
