"""Interactive dashboard for a deterministic synthetic e-commerce case study."""
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
