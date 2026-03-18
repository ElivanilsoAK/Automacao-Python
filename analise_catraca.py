import pandas as pd
import datetime

df = pd.read_csv('evento2026-03-18.csv', sep=';', encoding='utf-8', on_bad_lines='skip')
df.columns = df.columns.str.strip()

# Clean col names
col_map = {
    'Nome do Usuário': 'Nome', 
    'Acesso': 'Direcao',
    'Data Evento': 'Data',
    'Status': 'Status'
}
for c in df.columns:
    if c in col_map:
        df.rename(columns={c: col_map[c]}, inplace=True)
        
df = df[df['Status'].astype(str).str.contains('Liberado', case=False, na=False)]
df = df.dropna(subset=['Data', 'Nome', 'Direcao'])
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df = df.dropna(subset=['Data']).sort_values(['Nome', 'Data'])

# Normalize Direction
df['Dir'] = df['Direcao'].str.lower().str.normalize('NFKD').str.encode('ascii', 'ignore').str.decode('utf-8').str.strip()
df['Is_In'] = df['Dir'].str.contains('entr')

print("=== ANALISE DE ANOMALIAS ===")
grouped = df.groupby('Nome')

only_out = []
multi_hits = []
weird_seq = []

for name, group in grouped:
    group = group.sort_values('Data')
    
    # 1. Saiu sem entrar
    first_event = group.iloc[0]
    if not first_event['Is_In'] and len(group) == 1:
        only_out.append(name)
        
    # Sair antes de entrar (primeiro evento do dia foi saida)
    if not first_event['Is_In'] and len(group) > 1:
        weird_seq.append(name)
        
    # 2. Multi-hits em < 1 min
    group['diff'] = group['Data'].diff().dt.total_seconds()
    if (group['diff'] < 60).any():
        multi_hits.append(name)

print(f"Total de pessoas únicas no dia: {len(grouped)}")
print(f"1. Pessoas que SÓ tiveram registro de SAÍDA: {len(only_out)}")
print(f"2. Pessoas cujo PRIMEIRO registro do dia foi SAÍDA: {len(weird_seq)}")
print(f"3. Pessoas com múltiplas batidas em < 60s: {len(multi_hits)}")

print('\nExemplo de só saída:', only_out[:3])
print('Exemplo de multi-hits:', multi_hits[:3])

print("\nConclusão sobre 'Último Status':")
