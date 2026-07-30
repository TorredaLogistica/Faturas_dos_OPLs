from pathlib import Path
import io, os, unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import requests
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title='Indicador de Faturas',page_icon='📊',layout='wide')
BASE=Path(__file__).resolve().parent / 'Base Fatura.xlsx'
PALETA=['#E30613','#2188FF','#172B85','#FF9D76','#6A5ACD','#22A699']

# Compatibilidade com GitHub/Streamlit Cloud.
# Prioridade: arquivo no repositório, Secret do Streamlit e variável de ambiente.
try:
    BASE_URL=str(st.secrets.get('BASE_FATURA_URL','')).strip()
except Exception:
    BASE_URL=''
BASE_URL=BASE_URL or os.getenv('BASE_FATURA_URL','').strip()

def normalizar_url_github(url):
    url=str(url).strip()
    if 'github.com' in url and '/blob/' in url:
        url=url.replace('https://github.com/','https://raw.githubusercontent.com/').replace('/blob/','/')
    return url

def sem_acento(v): return ''.join(c for c in unicodedata.normalize('NFKD',str(v)) if not unicodedata.combining(c))
def empresa(v):
    s=sem_acento(v).strip().upper(); return {'CLARO MOVEL':'CLARO MÓVEL','CLARO FIXO':'CLARO FIXO','CLARO TV':'CLARO TV','EMBRATEL':'EMBRATEL','NET':'NET'}.get(s,str(v).strip())
def unidade(v): return 'Rio de Janeiro' if sem_acento(v).strip().upper()=='RIO DE JANEIRO' else str(v).strip().title()
@st.cache_data(ttl=3600,show_spinner='Carregando a base...')
def carregar(mtime=0):
    if BASE.exists():
        fonte=BASE
    elif BASE_URL:
        url=normalizar_url_github(BASE_URL)
        resposta=requests.get(url,timeout=120)
        resposta.raise_for_status()
        fonte=io.BytesIO(resposta.content)
    else:
        raise FileNotFoundError(
            'Base Fatura.xlsx não encontrada. Inclua o arquivo no mesmo diretório do app '
            'ou configure BASE_FATURA_URL nos Secrets do Streamlit.'
        )
    d=pd.read_excel(fonte,sheet_name='Fatura',engine='openpyxl')
    d['Data']=pd.to_datetime(d['MÊS/ANO'],errors='coerce'); d['Custo']=pd.to_numeric(d['Custo'],errors='coerce').fillna(0); d['Volume']=pd.to_numeric(d['Volume'],errors='coerce').fillna(0)
    d['Empresa']=d['Empresa'].map(empresa); d['Unidade']=d['Unidade'].map(unidade)
    for c in ['Tipo_Local','UF','Operador_logístico','Tipo_Despesa']: d[c]=d[c].astype(str).str.strip().str.upper()
    return d.dropna(subset=['Data'])
def br_num(v,n=2): return f'{v:,.{n}f}'.replace(',','X').replace('.',',').replace('X','.')
def brl(v): return f'R$ {br_num(v,2)}'
def graf_brl(v): return '' if pd.isna(v) else brl(float(v))
def ticks_brl(vals,n=6):
    x=pd.Series(vals).dropna().astype(float); vmax=max(x.max()*1.15,1); t=np.linspace(0,vmax,n); return t,[brl(v) for v in t]
def card(t,v,s,c): st.markdown(f'<div class="kpi" style="border-top-color:{c}"><div class="kt">{t}</div><div class="kv">{v}</div><div class="ks">{s}</div></div>',unsafe_allow_html=True)
def montar_tab(base,tipo):
    z=base[base.Tipo_Despesa.eq(tipo)].groupby('Unidade',as_index=False).agg(Custo=('Custo','sum'),**{'Volume (Real)':('Volume','sum')})
    col='Unitário IN' if tipo=='VOL IN' else 'Unitário OUT'; z[col]=np.where(z['Volume (Real)'].ne(0),z.Custo/z['Volume (Real)'],np.nan); z=z.sort_values('Unidade')
    c=z.Custo.sum(); v=z['Volume (Real)'].sum(); return pd.concat([z,pd.DataFrame([{'Unidade':'Total','Custo':c,'Volume (Real)':v,col:c/v if v else np.nan}])],ignore_index=True)
def tabela_unitaria_formatada(tab, coluna_unitaria):
    exib=tab.copy()
    exib['Custo']=exib['Custo'].map(brl)
    exib['Volume (Real)']=exib['Volume (Real)'].map(lambda v: br_num(v,0))
    exib[coluna_unitaria]=exib[coluna_unitaria].map(lambda v: '' if pd.isna(v) else brl(v))
    def destacar_total(linha):
        if linha['Unidade']=='Total':
            return ['font-weight:800; background-color:#D9EAF7; color:#172033']*len(linha)
        return ['']*len(linha)
    return exib.style.apply(destacar_total,axis=1).set_properties(**{'text-align':'right'}).set_properties(subset=['Unidade'],**{'text-align':'left'})

def rotulo_milhoes(v):
    return f'{float(v)/1_000_000:.2f} Mi'.replace('.',',')

def rotulo_compacto(v):
    valor=float(v)
    if abs(valor) >= 1_000_000:
        return f'{valor/1_000_000:.2f} Mi'.replace('.',',')
    if abs(valor) >= 1_000:
        return f'{valor/1_000:.2f} Mil'.replace('.',',')
    return brl(valor)

def rotulo_compacto_reais(v):
    valor=float(v)
    if abs(valor) >= 1_000_000:
        return ('R$ ' + f'{valor/1_000_000:.2f}'.replace('.',',') + ' Mi')
    if abs(valor) >= 1_000:
        return ('R$ ' + f'{valor/1_000:.2f}'.replace('.',',') + ' Mil')
    return brl(valor)

def montar_custos_volumes(base):
    z=base.groupby(['Unidade','Tipo_Local'],as_index=False).agg(Custo=('Custo','sum'),Volume=('Volume','sum'))
    z['Custo unitário']=np.where(z['Volume'].ne(0),z['Custo']/z['Volume'],np.nan)
    z=z.sort_values(['Unidade','Tipo_Local']).reset_index(drop=True)
    custo=z['Custo'].sum(); volume=z['Volume'].sum()
    total=pd.DataFrame([{'Unidade':'Total','Tipo_Local':'','Custo':custo,'Volume':volume,'Custo unitário':np.nan}])
    return pd.concat([z,total],ignore_index=True)

def tabela_custos_volumes_formatada(tab):
    exib=tab.copy()
    exib['Custo']=exib['Custo'].map(brl)
    exib['Volume']=exib['Volume'].map(lambda v: br_num(v,2))
    exib['Custo unitário']=exib['Custo unitário'].map(lambda v: '' if pd.isna(v) else brl(v))
    def destacar(linha):
        return ['font-weight:800; background-color:#D9EAF7; color:#172033']*len(linha) if linha['Unidade']=='Total' else ['']*len(linha)
    return exib.style.apply(destacar,axis=1).set_properties(**{'text-align':'right'}).set_properties(subset=['Unidade','Tipo_Local'],**{'text-align':'left'})

def excel_completo(tin,tout,tcv):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as w:
        conjuntos=[('Custos e Volumes',tcv),('VOL IN',tin),('VOL OUT',tout)]
        for nome,t in conjuntos:
            t.to_excel(w,sheet_name=nome,index=False,startrow=2)
            ws=w.book[nome]; ws.sheet_view.showGridLines=False
            ws['A1']='Custos e volumes por unidade' if nome=='Custos e Volumes' else f'Detalhamento de custo unitário - {nome}'
            fim_col=5 if nome=='Custos e Volumes' else 4
            ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=fim_col)
            ws['A1'].font=Font(bold=True,color='FFFFFF',size=14); ws['A1'].fill=PatternFill('solid',fgColor='404B5A' if nome=='Custos e Volumes' else ('2188FF' if nome=='VOL IN' else '172B85'))
            for c in ws[3]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='404B5A'); c.alignment=Alignment(horizontal='center')
            for r in range(4,ws.max_row+1):
                if nome=='Custos e Volumes':
                    ws.cell(r,3).number_format='R$ #,##0.00;[Red](R$ #,##0.00);-'; ws.cell(r,4).number_format='#,##0.00;[Red](#,##0.00);-'; ws.cell(r,5).number_format='R$ #,##0.00;[Red](R$ #,##0.00);-'
                else:
                    ws.cell(r,2).number_format='R$ #,##0.00;[Red](R$ #,##0.00);-'; ws.cell(r,3).number_format='#,##0;[Red](#,##0);-'; ws.cell(r,4).number_format='R$ #,##0.00;[Red](R$ #,##0.00);-'
                if ws.cell(r,1).value=='Total':
                    for c in range(1,fim_col+1): ws.cell(r,c).font=Font(bold=True); ws.cell(r,c).fill=PatternFill('solid',fgColor='D9EAF7')
            widths=[24,16,20,18,20] if nome=='Custos e Volumes' else [24,18,18,18]
            for i,width in enumerate(widths,1): ws.column_dimensions[chr(64+i)].width=width
            ws.freeze_panes='A4'
    return b.getvalue()

def excel_tabs(tin,tout):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as w:
        for nome,t in [('VOL IN',tin),('VOL OUT',tout)]:
            t.to_excel(w,sheet_name=nome,index=False,startrow=2); ws=w.book[nome]; ws.sheet_view.showGridLines=False; ws['A1']=f'Detalhamento de custo unitário - {nome}'; ws.merge_cells('A1:D1'); ws['A1'].font=Font(bold=True,color='FFFFFF',size=14); ws['A1'].fill=PatternFill('solid',fgColor='2188FF' if nome=='VOL IN' else '172B85')
            for c in ws[3]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='404B5A'); c.alignment=Alignment(horizontal='center')
            for r in range(4,ws.max_row+1):
                ws.cell(r,2).number_format='#,##0.00;[Red](#,##0.00);-'; ws.cell(r,3).number_format='#,##0;[Red](#,##0);-'; ws.cell(r,4).number_format='#,##0.00;[Red](#,##0.00);-'
                if ws.cell(r,1).value=='Total':
                    for c in range(1,5): ws.cell(r,c).font=Font(bold=True); ws.cell(r,c).fill=PatternFill('solid',fgColor='D9EAF7')
            for col,width in zip('ABCD',[24,18,18,18]): ws.column_dimensions[col].width=width
            ws.freeze_panes='A4'
    return b.getvalue()

st.markdown('''<style>.block-container{padding-top:1rem;max-width:1800px}.kpi{background:white;border:1px solid #d9dee7;border-top:5px solid;border-radius:14px;padding:18px 14px;text-align:center;min-height:128px;box-shadow:0 2px 8px #0000000d}.kt{font-size:.82rem;font-weight:800;color:#4d5766;text-transform:uppercase;white-space:nowrap}.kv{font-size:clamp(1.35rem,2.1vw,2rem);font-weight:900;color:#172033;margin-top:12px;white-space:nowrap;line-height:1.1}.ks{font-size:.82rem;color:#6b7280;margin-top:12px}.opl-card{background:#F6C2AA;border:1px solid #4A4A4A;border-radius:7px;padding:20px 14px;text-align:center;min-height:116px;box-shadow:0 1px 3px #00000012}.opl-value{font-size:clamp(1.65rem,2.6vw,2.35rem);font-weight:900;color:#202428;line-height:1.12;white-space:nowrap}.opl-label{font-size:1rem;color:#616161;margin-top:10px}.stTabs [data-baseweb="tab"]{font-weight:800}</style>''',unsafe_allow_html=True)
try: df=carregar(BASE.stat().st_mtime if BASE.exists() else 0)
except Exception as e: st.error(e); st.stop()
st.title('📊 Indicador de Custos de Faturas — CDs e EAs'); st.caption(f'Base atualizada até {df.Data.max():%m/%Y} • {len(df):,.0f} registros'.replace(',','.'))
with st.sidebar:
    st.header('Filtros'); datas=sorted(df.Data.unique()); per=st.multiselect('Mês/Ano',datas,default=[max(datas)],format_func=lambda x:pd.Timestamp(x).strftime('%m/%Y'))
    cas=df[df.Data.isin(per)] if per else df
    sels={}
    for col,rot in [('Tipo_Local','Tipo de local'),('Empresa','Empresa'),('UF','UF'),('Unidade','Unidade'),('Operador_logístico','Operador logístico'),('Tipo_Despesa','Tipo de despesa')]:
        opts=sorted(cas[col].dropna().unique().tolist()); sels[col]=st.multiselect(rot,opts,placeholder='Todos');
        if sels[col]: cas=cas[cas[col].isin(sels[col])]
f=df.copy()
if per:f=f[f.Data.isin(per)]
for c,v in sels.items():
    if v:f=f[f[c].isin(v)]
if f.empty: st.warning('Nenhum registro encontrado.'); st.stop()
ct=f.Custo.sum(); cd=f.loc[f.Tipo_Local.eq('CD'),'Custo'].sum(); ea=f.loc[f.Tipo_Local.eq('EA'),'Custo'].sum(); vol=f.Volume.sum(); bi=f[f.Tipo_Despesa.eq('VOL IN')]; bo=f[f.Tipo_Despesa.eq('VOL OUT')]; ui=bi.Custo.sum()/bi.Volume.sum() if bi.Volume.sum() else np.nan; uo=bo.Custo.sum()/bo.Volume.sum() if bo.Volume.sum() else np.nan
fim=pd.Timestamp(max(per)) if per else df.Data.max(); ini=fim-pd.DateOffset(months=11); hist=df[df.Data.between(ini,fim)].copy()
for c,v in sels.items():
    if v:hist=hist[hist[c].isin(v)]
prev=df[df.Data.eq(fim-pd.DateOffset(months=1))]
for c,v in sels.items():
    if v:prev=prev[prev[c].isin(v)]
var=ct/prev.Custo.sum()-1 if prev.Custo.sum() else np.nan
abas=st.tabs(['Visão executiva','Custos e volumes'])
with abas[0]:
    a=st.columns(3)
    with a[0]:card('Custo CD + EA',brl(ct),'Fatura total',PALETA[0])
    with a[1]:card('Custo CD',brl(cd),'Centros de distribuição',PALETA[1])
    with a[2]:card('Custo EA',brl(ea),'Estações avançadas',PALETA[2])
    st.write(''); a=st.columns(3)
    with a[0]:card('Volume total',br_num(vol,0),'Itens/movimentos',PALETA[3])
    with a[1]:card('Unitário IN | OUT',('—' if np.isnan(ui) else brl(ui))+' | '+('—' if np.isnan(uo) else brl(uo)),'Custo ÷ volume por fluxo',PALETA[4])
    with a[2]:card('Variação mensal','—' if np.isnan(var) else f'{var:+.1%}'.replace('.',','),'Mesmo recorte vs mês anterior',PALETA[5])
    st.markdown('<div style="height:14px"></div>',unsafe_allow_html=True)
    opl_cards=st.columns(3)
    cores_opl={'CELISTICS':'#000000','POSTALGOW':'#000000','TPC':'#000000'}
    for coluna,nome in zip(opl_cards,['CELISTICS','POSTALGOW','TPC']):
        valor_opl=f.loc[f['Operador_logístico'].eq(nome),'Custo'].sum()
        with coluna:
            card(f'CUSTO {nome}',brl(valor_opl),f'{nome} (Fatura)',cores_opl[nome])
    st.markdown('<div style="height:12px"></div>',unsafe_allow_html=True)
    st.markdown('<div style="font-size:18px;font-weight:800;color:#172033;margin:2px 0 4px 0;">Evolução mensal do custo — últimos 12 meses</div>',unsafe_allow_html=True)
    m=hist.groupby('Data',as_index=False).Custo.sum().sort_values('Data'); m['Rotulo']=m.Custo.map(rotulo_milhoes); m['Moeda']=m.Custo.map(graf_brl)
    fig=px.line(m,x='Data',y='Custo',markers=False,color_discrete_sequence=[PALETA[1]],custom_data=['Moeda']); fig.update_traces(line_width=3,hovertemplate='%{x|%m/%Y}<br>%{customdata[0]}<extra></extra>')
    for i,r in m.reset_index(drop=True).iterrows():
        fig.add_annotation(x=r.Data,y=r.Custo,text=r.Rotulo,showarrow=False,yshift=16 if i%2==0 else -18,font=dict(size=11,color='#666666',family='Arial Black'),bgcolor='rgba(255,255,255,.76)',borderpad=1)
    ymin=float(m.Custo.min()); ymax=float(m.Custo.max()); faixa=max(ymax-ymin,ymax*.03)
    # Exibe todos os 12 meses no padrão MM/AAAA, igual aos demais gráficos do painel.
    fig.update_xaxes(tickmode='array',tickvals=m['Data'],ticktext=m['Data'].dt.strftime('%m/%Y'),tickangle=0,showgrid=True,gridcolor='#D6D6D6',griddash='dot',showline=False,range=[ini-pd.Timedelta(days=14),fim+pd.Timedelta(days=14)])
    fig.update_yaxes(range=[ymin-faixa*.28,ymax+faixa*.30],showticklabels=False,showgrid=False,zeroline=False,title='')
    fig.update_layout(height=350,xaxis_title='',hovermode='x unified',margin=dict(l=10,r=20,t=10,b=35),plot_bgcolor='white',showlegend=False)
    st.plotly_chart(fig,use_container_width=True)

    # Variação percentual mês a mês: utiliza 13 meses para calcular corretamente a primeira variação exibida.
    ini_variacao=ini-pd.DateOffset(months=1)
    hist_variacao=df[df.Data.between(ini_variacao,fim)].copy()
    for c,v in sels.items():
        if v:
            hist_variacao=hist_variacao[hist_variacao[c].isin(v)]
    variacao_mensal=hist_variacao.groupby('Data',as_index=False).Custo.sum().sort_values('Data')
    variacao_mensal['Variacao']=variacao_mensal['Custo'].pct_change()*100
    variacao_mensal=variacao_mensal[variacao_mensal['Data'].between(ini,fim)].copy()
    variacao_mensal['Rotulo']=variacao_mensal['Variacao'].map(lambda v: '' if pd.isna(v) else f'{v:+.1f}%'.replace('.',','))
    variacao_mensal['Cor']=np.where(variacao_mensal['Variacao'].ge(0),'Positiva','Negativa')
    fig_var=px.bar(variacao_mensal,x='Data',y='Variacao',color='Cor',text='Rotulo',title='Variação mensal do custo — últimos 12 meses',color_discrete_map={'Positiva':'#22A699','Negativa':'#E30613'},custom_data=['Rotulo'])
    fig_var.update_traces(textposition='outside',cliponaxis=False,textfont=dict(size=12,family='Arial Black'),hovertemplate='%{x|%m/%Y}<br>Variação: %{customdata[0]}<extra></extra>')
    maior_var=max(float(variacao_mensal['Variacao'].abs().max()) if variacao_mensal['Variacao'].notna().any() else 1,1)
    fig_var.update_xaxes(tickmode='array',tickvals=variacao_mensal['Data'],ticktext=variacao_mensal['Data'].dt.strftime('%m/%Y'),tickangle=0,showgrid=False,range=[ini-pd.Timedelta(days=14),fim+pd.Timedelta(days=14)])
    fig_var.update_yaxes(title='Variação (%)',ticksuffix='%',range=[-maior_var*1.35,maior_var*1.35],zeroline=True,zerolinecolor='#7A8491',zerolinewidth=1,gridcolor='#E4E8EF')
    fig_var.update_layout(height=390,xaxis_title='',showlegend=False,plot_bgcolor='white',margin=dict(l=20,r=25,t=60,b=40),font=dict(size=12),title=dict(font=dict(size=18,color='#172033'),x=0.0,xanchor='left'))
    st.plotly_chart(fig_var,use_container_width=True)
    uh=hist[hist.Tipo_Despesa.isin(['VOL IN','VOL OUT'])].groupby(['Data','Tipo_Despesa'],as_index=False).agg(Custo=('Custo','sum'),Volume=('Volume','sum')); uh['Unitario']=np.where(uh.Volume.ne(0),uh.Custo/uh.Volume,np.nan); uh['Rotulo']=uh.Unitario.map(rotulo_compacto)
    fig=px.line(uh,x='Data',y='Unitario',color='Tipo_Despesa',markers=True,custom_data=['Rotulo'],color_discrete_map={'VOL IN':PALETA[1],'VOL OUT':PALETA[2]},title='Evolução mensal do custo unitário IN e OUT — últimos 12 meses'); fig.update_traces(line_width=4,marker_size=10,hovertemplate='%{x|%m/%Y}<br>%{fullData.name}: %{customdata[0]}<extra></extra>')
    for _,r in uh.iterrows(): fig.add_annotation(x=r.Data,y=r.Unitario,text=r.Rotulo,showarrow=False,yshift=22 if r.Tipo_Despesa=='VOL OUT' else 23,font=dict(size=11,color=PALETA[2] if r.Tipo_Despesa=='VOL OUT' else PALETA[1],family='Arial Black'),bgcolor='rgba(255,255,255,.88)',borderpad=2)
    tv,tt=ticks_brl(uh.Unitario); fig.update_xaxes(dtick='M1',tickformat='%m/%Y',tickangle=0); fig.update_yaxes(tickvals=tv,ticktext=tt,rangemode='tozero'); fig.update_layout(height=520,xaxis_title='',yaxis_title='Custo unitário (R$)',hovermode='x unified',margin=dict(l=20,r=30,t=65,b=45),font=dict(size=13),legend=dict(font=dict(size=13)),title=dict(font=dict(size=18,color='#172033'),x=0.0,xanchor='left')); st.plotly_chart(fig,use_container_width=True)
    fluxo=hist[hist.Tipo_Despesa.isin(['VOL IN','VOL IN E-COMMERCE','VOL OUT','VOL OUT E-COMMERCE'])].copy(); fluxo['Fluxo']=np.where(fluxo.Tipo_Despesa.str.startswith('VOL IN'),'IN','OUT'); fluxo=fluxo.groupby(['Data','Fluxo'],as_index=False).Custo.sum(); fluxo['Rotulo']=fluxo.Custo.map(rotulo_compacto)
    fig=px.line(fluxo,x='Data',y='Custo',color='Fluxo',markers=True,custom_data=['Rotulo'],color_discrete_map={'IN':PALETA[1],'OUT':PALETA[2]},title='Evolução do custo por despesas IN e OUT — últimos 12 meses')
    fig.update_traces(line_width=4,marker_size=10,hovertemplate='%{x|%m/%Y}<br>%{fullData.name}: %{customdata[0]}<extra></extra>')
    for _,r in fluxo.iterrows():
        fig.add_annotation(x=r.Data,y=r.Custo,text=r.Rotulo,showarrow=False,yshift=22 if r.Fluxo=='OUT' else 20,font=dict(size=11,color=PALETA[2] if r.Fluxo=='OUT' else PALETA[1],family='Arial Black'),bgcolor='rgba(255,255,255,.88)',borderpad=2)
    tv,tt=ticks_brl(fluxo.Custo); fig.update_xaxes(dtick='M1',tickformat='%m/%Y',tickangle=0); fig.update_yaxes(tickvals=tv,ticktext=tt,rangemode='tozero'); fig.update_layout(height=540,xaxis_title='',yaxis_title='Custo (R$)',hovermode='x unified',margin=dict(l=20,r=35,t=65,b=35),font=dict(size=13),legend=dict(font=dict(size=13)),title=dict(font=dict(size=18,color='#172033'),x=0.0,xanchor='left')); st.plotly_chart(fig,use_container_width=True)
    c1,c2=st.columns(2)
    with c1:
        u=f.groupby(['Unidade','Tipo_Local'],as_index=False).Custo.sum().sort_values('Custo',ascending=False).head(12)
        u['Rotulo']=u.Custo.map(rotulo_compacto_reais)
        totais_unidade=u.groupby('Unidade')['Custo'].sum()
        # Em barras estreitas, os dois valores ficam após a barra para evitar truncamento.
        unidades_externas=set(totais_unidade.index)
        u['Rotulo_barra']=np.where(u['Unidade'].isin(unidades_externas),'',u['Rotulo'])
        fig=px.bar(u,x='Custo',y='Unidade',color='Tipo_Local',orientation='h',text='Rotulo_barra',title=None,custom_data=['Custo'],color_discrete_map={'CD':'#0D73C9','EA':'#79BFF2'})
        for trace in fig.data:
            if trace.name=='CD':
                trace.update(textposition='inside',insidetextanchor='end',textfont=dict(size=11,color='white',family='Arial Black'))
            else:
                trace.update(textposition='outside',textfont=dict(size=11,color='#79BFF2',family='Arial Black'))
            trace.update(cliponaxis=False,hovertemplate='%{y}<br>R$ %{customdata[0]:,.2f}<extra></extra>')
        # Rótulos externos por segmento, com a cor correspondente a CD e EA.
        for unidade_nome in unidades_externas:
            linhas=u[u['Unidade'].eq(unidade_nome)]
            total_unidade=float(linhas['Custo'].sum())
            linha_cd=linhas[linhas['Tipo_Local'].eq('CD')]
            linha_ea=linhas[linhas['Tipo_Local'].eq('EA')]
            deslocamento=7
            if not linha_cd.empty:
                rot_cd=linha_cd.iloc[0]['Rotulo']
                fig.add_annotation(x=total_unidade,y=unidade_nome,text=rot_cd,showarrow=False,xanchor='left',xshift=deslocamento,font=dict(size=11,color='#0D73C9',family='Arial Black'),bgcolor='rgba(255,255,255,.92)',borderpad=1)
                deslocamento += max(68, len(rot_cd)*5)
            if not linha_ea.empty:
                rot_ea=linha_ea.iloc[0]['Rotulo']
                fig.add_annotation(x=total_unidade,y=unidade_nome,text='/ '+rot_ea,showarrow=False,xanchor='left',xshift=deslocamento,font=dict(size=11,color='#79BFF2',family='Arial Black'),bgcolor='rgba(255,255,255,.92)',borderpad=1)
        fig.update_layout(yaxis={'categoryorder':'total ascending','tickfont':dict(size=13),'domain':[0.0,1.0]},xaxis={'tickfont':dict(size=12),'title_font':dict(size=14)},height=540,margin=dict(l=15,r=285,t=25,b=45),showlegend=False,barmode='stack',uniformtext_minsize=10,uniformtext_mode='show')
        st.markdown('''
        <div style="display:flex;align-items:center;justify-content:space-between;gap:16px;margin:2px 0 4px 0;">
          <div style="font-size:18px;font-weight:800;color:#172033;white-space:nowrap;">Ranking de custo por unidade</div>
          <div style="display:flex;align-items:center;gap:14px;font-size:13px;color:#4D5766;white-space:nowrap;padding-right:8px;">
            <span style="font-weight:700;">Tipo_Local:</span>
            <span style="display:flex;align-items:center;gap:6px;"><span style="width:13px;height:13px;background:#0D73C9;display:inline-block;border-radius:2px;"></span>CD</span>
            <span style="display:flex;align-items:center;gap:6px;"><span style="width:13px;height:13px;background:#79BFF2;display:inline-block;border-radius:2px;"></span>EA</span>
          </div>
        </div>
        ''',unsafe_allow_html=True)
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        d=f.groupby('Tipo_Despesa',as_index=False).Custo.sum().sort_values('Custo',ascending=False).head(12)
        d['Rotulo']=d.Custo.map(rotulo_compacto_reais)
        d['Rotulo_barra']=np.where(d.Custo.ge(0),d.Rotulo,'')
        fig=px.bar(d,x='Custo',y='Tipo_Despesa',orientation='h',text='Rotulo_barra',title='Composição do custo por despesa',color_discrete_sequence=[PALETA[0]],custom_data=['Custo'])
        fig.update_traces(textposition='outside',textfont=dict(size=12,family='Arial Black'),cliponaxis=False,hovertemplate='%{y}<br>R$ %{customdata[0]:,.2f}<extra></extra>')
        maior_positivo=max(float(d.loc[d.Custo.ge(0),'Custo'].max()),1)
        posicao_rotulo_negativo=maior_positivo*0.025
        for _,linha in d[d.Custo.lt(0)].iterrows():
            fig.add_annotation(x=posicao_rotulo_negativo,y=linha.Tipo_Despesa,text=linha.Rotulo,showarrow=False,xanchor='left',font=dict(size=12,color='#E30613',family='Arial Black'),bgcolor='rgba(255,255,255,.96)',borderpad=2)
        maior_abs=max(float(d.Custo.abs().max()),1)
        fig.update_xaxes(range=[min(float(d.Custo.min())*1.35,-maior_abs*.07),float(d.Custo.max())*1.22],zeroline=True,zerolinecolor='#AAB2BD',zerolinewidth=1)
        fig.update_layout(yaxis={'categoryorder':'total ascending','tickfont':dict(size=13),'title_font':dict(size=14)},xaxis={'tickfont':dict(size=12),'title_font':dict(size=14)},height=500,margin=dict(l=15,r=190,t=85,b=40),title=dict(font=dict(size=18)))
        st.plotly_chart(fig,use_container_width=True)

    # Ranking de custo por empresa e Tipo_Local, no mesmo padrão do ranking por unidade.
    st.markdown('<div style="height:18px"></div>',unsafe_allow_html=True)
    ranking_empresa=f.groupby(['Empresa','Tipo_Local'],as_index=False).Custo.sum()
    totais_empresa=ranking_empresa.groupby('Empresa',as_index=False).Custo.sum().sort_values('Custo',ascending=False).head(12)
    empresas_top=totais_empresa['Empresa'].tolist()
    ranking_empresa=ranking_empresa[ranking_empresa['Empresa'].isin(empresas_top)].copy()
    ranking_empresa['Rotulo']=ranking_empresa['Custo'].map(rotulo_compacto_reais)
    ranking_empresa['Rotulo_barra']=''
    fig_empresa=px.bar(ranking_empresa,x='Custo',y='Empresa',color='Tipo_Local',orientation='h',text='Rotulo_barra',title=None,custom_data=['Custo'],color_discrete_map={'CD':'#16856F','EA':'#78C9A8'})
    fig_empresa.update_traces(cliponaxis=False,hovertemplate='%{y}<br>%{fullData.name}: R$ %{customdata[0]:,.2f}<extra></extra>')
    for empresa_nome in empresas_top:
        linhas_empresa=ranking_empresa[ranking_empresa['Empresa'].eq(empresa_nome)]
        total_empresa=float(linhas_empresa['Custo'].sum())
        linha_cd=linhas_empresa[linhas_empresa['Tipo_Local'].eq('CD')]
        linha_ea=linhas_empresa[linhas_empresa['Tipo_Local'].eq('EA')]
        deslocamento=7
        if not linha_cd.empty:
            rot_cd=linha_cd.iloc[0]['Rotulo']
            fig_empresa.add_annotation(x=total_empresa,y=empresa_nome,text=rot_cd,showarrow=False,xanchor='left',xshift=deslocamento,font=dict(size=11,color='#16856F',family='Arial Black'),bgcolor='rgba(255,255,255,.92)',borderpad=1)
            deslocamento += max(68,len(rot_cd)*5)
        if not linha_ea.empty:
            rot_ea=linha_ea.iloc[0]['Rotulo']
            fig_empresa.add_annotation(x=total_empresa,y=empresa_nome,text='/ '+rot_ea,showarrow=False,xanchor='left',xshift=deslocamento,font=dict(size=11,color='#78C9A8',family='Arial Black'),bgcolor='rgba(255,255,255,.92)',borderpad=1)
    fig_empresa.update_layout(yaxis={'categoryorder':'array','categoryarray':list(reversed(empresas_top)),'tickfont':dict(size=13),'title_font':dict(size=14)},xaxis={'tickfont':dict(size=12),'title_font':dict(size=14)},height=max(500,90+len(empresas_top)*58),margin=dict(l=15,r=310,t=25,b=45),showlegend=False,barmode='stack',plot_bgcolor='white')
    legenda_empresa = '''<div style="display:flex;align-items:center;justify-content:space-between;gap:16px;margin:2px 0 4px 0;"><div style="font-size:18px;font-weight:800;color:#172033;white-space:nowrap;">Ranking de custo por empresa</div><div style="display:flex;align-items:center;gap:14px;font-size:13px;color:#4D5766;white-space:nowrap;padding-right:8px;"><span style="font-weight:700;">Tipo_Local:</span><span style="display:flex;align-items:center;gap:6px;"><span style="width:13px;height:13px;background:#16856F;display:inline-block;border-radius:2px;"></span>CD</span><span style="display:flex;align-items:center;gap:6px;"><span style="width:13px;height:13px;background:#78C9A8;display:inline-block;border-radius:2px;"></span>EA</span></div></div>'''
    st.markdown(legenda_empresa,unsafe_allow_html=True)
    st.plotly_chart(fig_empresa,use_container_width=True)
with abas[1]:
    st.subheader('Custos e volumes por unidade')
    tcv=montar_custos_volumes(f)
    st.dataframe(tabela_custos_volumes_formatada(tcv),use_container_width=True,hide_index=True,height=520)

    st.divider()
    st.subheader('Detalhamento de custos unitários IN e OUT')
    st.caption('Cálculo: soma do custo ÷ soma do volume real. As tabelas respeitam todos os filtros selecionados.')
    tin=montar_tab(f,'VOL IN'); tout=montar_tab(f,'VOL OUT')
    c1,c2=st.columns(2)
    with c1:
        st.markdown('**Tipo_Despesa: VOL IN**')
        st.dataframe(tabela_unitaria_formatada(tin,'Unitário IN'),use_container_width=True,hide_index=True,height=420)
    with c2:
        st.markdown('**Tipo_Despesa: VOL OUT**')
        st.dataframe(tabela_unitaria_formatada(tout,'Unitário OUT'),use_container_width=True,hide_index=True,height=420)

    arquivo_excel=excel_completo(tin,tout,tcv)
    st.download_button('📥 Baixar tabelas em Excel',arquivo_excel,'Custos_Volumes_e_Unitarios.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
