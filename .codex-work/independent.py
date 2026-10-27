import csv, json, math, statistics
from pathlib import Path

root = Path(__file__).resolve().parent.parent
results = json.loads((root/'.codex-work/product-results.json').read_text(encoding='utf-8-sig'))
daily = {}
with (root/'exports/T_Commandes.csv').open(encoding='utf-8-sig', newline='') as f:
    orders = list(csv.DictReader(f))
for row in orders:
    if row['Statut'] != 'Annulee':
        key = row['Produit'], row['Date'][:10]
        daily[key] = daily.get(key, 0) + float(row['Quantite'])
z = statistics.NormalDist().inv_cdf(.95)
for data in results:
    product = data['product']
    ds = [v for (p,d),v in daily.items() if p == product]
    mu, sigma = statistics.mean(ds), statistics.stdev(ds)
    assert math.isclose(mu,data['mu'],rel_tol=1e-12)
    assert math.isclose(sigma,data['sigma'],rel_tol=1e-12)
    protection = mu*(data['L']+1)
    sd = sigma*math.sqrt(data['L']+1)
    lot = max(1,mu,math.sqrt(2*mu*data['Ccmd']/data['Cpos']))
    s,q = math.ceil(protection+z*sd),math.ceil(lot)
    s0 = math.ceil(s+q/2)
    assert (s,q,s0) == (data['s'],data['Q'],data['S0'])
    expected = [(math.ceil(protection+(z+.75)*sd),math.ceil(lot*1.35)),
                (math.ceil(protection+max(0,z-.75)*sd),math.ceil(max(mu,lot*.7))), (s,q)]
    for policy,(s,q) in zip(data['policies'],expected):
        assert policy[2:] == [s0,s,q]
    state = float(20261023)
    def uniform():
        global state
        state = (state*16807) % 2147483647
        return state/2147483647
    demands = [[max(0,math.floor(mu+sigma*math.sqrt(-2*math.log(uniform()))*math.cos(6.28318530717959*uniform())+.5)) for t in range(365)] for r in range(30)]
    for policy,stats in zip(data['policies'],data['results']):
        costs,stocks,counts,services,losts = [],[],[],[],[]
        for demand in demands:
            arrivals = [0]*(365+int(data['L']))
            stock,transit,total_stock,lost,ordered = int(s0),0,0,0,0
            for t,d in enumerate(demand):
                stock += arrivals[t]; transit -= arrivals[t]
                sale = min(stock,d); stock -= sale; lost += d-sale
                total_stock += stock
                if stock+transit <= policy[3]:
                    arrivals[t+int(data['L'])] += int(policy[4])
                    transit += int(policy[4]); ordered += 1
            # No lost units for the four calibrated products in this fixed sample.
            assert lost == 0
            costs.append(total_stock*data['Cpos'] + ordered*data['Ccmd'])
            stocks.append(total_stock/365); counts.append(ordered)
            services.append(1-lost/sum(demand)); losts.append(lost)
        for actual,want in [(stats[1],statistics.mean(costs)),(stats[2],statistics.stdev(costs)),(stats[5],statistics.mean(services)),(stats[9],statistics.mean(losts)),(stats[10],statistics.mean(counts)),(stats[11],statistics.mean(stocks))]:
            assert math.isclose(actual,want,abs_tol=1e-8,rel_tol=1e-10),(product,policy,actual,want)
print('PASS: historique des 4 produits, 12 suggestions et S0, simulation independante de 360 repetitions, cout/service/stock/commandes.')
print('Commandes historiques:', len(orders))
