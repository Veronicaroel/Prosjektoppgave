import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter

# a) Les inn Excel-filen og lagre data i lister 
filnavn = "support_uke_24.xlsx"
data = pd.read_excel(filnavn, engine="openpyxl")

# Lagre data i separate arrays
u_dag = data.iloc[:, 0].tolist()       # Ukedag
kl_slett = data.iloc[:, 1].tolist()    # Klokkeslett
varighet = data.iloc[:, 2].tolist()    # Samtalens varighet
score = data.iloc[:, 3].tolist()       # Kundens tilfredshet

# b) Visualiser med søylediagram

ukedag_teller = data.iloc[:, 0].value_counts() # Finner forekomst av ukedager


ukedag_counter = Counter(u_dag) # Bruker Counter for å telle ukedager

# Skriv ut antall henvendelser per ukedag
print("\nAntall henvendelser per ukedag:")
print(ukedag_teller)

# Formatering av diagram
plt.bar(ukedag_teller.index, ukedag_teller.values, color='blue')
plt.ylabel("Henvendelser")
plt.title("Antall supporthenvendelser per ukedag")
plt.show()

# c) Finn korteste og lengste samtaletid

# Funksjon for å konvertere "HH:MM:SS" til minutter (desimaltall)
def tid_til_minutter(tid):
    hh, mm, ss = map(int, tid.split(":"))  # Del opp i timer, minutter og sekunder
    return hh * 60 + mm + ss / 60          # Konverter til minutter

# Konverter varighet til minutter
varighet_minutter = [tid_til_minutter(tid) for tid in varighet]

min_varighet = min(varighet_minutter)
max_varighet = max(varighet_minutter)

# Skriv ut resultatet
print(f"\nKorteste samtale var {min_varighet:.2f} minutter.")
print(f"\nLengste samtale var {max_varighet:.2f} minutter.")

# d) Beregn gjennomsnittlig samtaletid
gjennomsnitt_varighet = sum(varighet_minutter) / len(varighet_minutter)

print(f"\nGjennomsnittlig samtaletid: {gjennomsnitt_varighet:.2f} minutter.")

# e) Antall henvendelser supportavdelingen mottok for hver av tidsrommene

# Hent klokkeslett-kolonnen som liste
kl_slett = data.iloc[:, 1].tolist()  

# Funksjon for å finne riktig tidsintervall for et klokkeslett
def finn_tidspunkt_intervall(tid):
    hh, mm = map(int, tid.split(":")[:2])  # Hent timer og minutter
    if 8 <= hh < 10:
        return "08-10"
    elif 10 <= hh < 12:
        return "10-12"
    elif 12 <= hh < 14:
        return "12-14"
    elif 14 <= hh < 16:
        return "14-16"
    else:
        return "Utenfor arbeidstid"

# Lag en liste med intervaller
tid_intervaller = [finn_tidspunkt_intervall(tid) for tid in kl_slett]

# Tell forekomstene av hver tidsperiode
tid_telling = pd.Series(tid_intervaller).value_counts().reindex(["08-10", "10-12", "12-14", "14-16"], fill_value=0)

# Skriv ut resultatet
print("\nAntall henvendelser per tidsbolk:")
print(tid_telling)

# e) Visualiser med sektordiagram (kakediagram)
plt.pie(tid_telling, labels=tid_telling.index, autopct='%1.1f%%', colors=["blue", "yellow", "green", "red"])
plt.title("Antall henvendelser per tidsbolk (Uke 24)")
plt.show() 


# f) Supportavdelingens Net Promoter Score

# Henter kundetilfredshets-kolonnen
score = data.iloc[:, 3].dropna().tolist()  # Fjerner manglende verdier

# Klassifiser kunder i henhold til tilfredshet
antall_kunder = len(score)
negative = sum(1 for s in score if 1 <= s <= 6)
nøytrale = sum(1 for s in score if 7 <= s <= 8)
positive = sum(1 for s in score if 9 <= s <= 10)

# Bereger prosentandel
prosent_positive = (positive / antall_kunder) * 100
prosent_negative = (negative / antall_kunder) * 100
prosent_nøytrale = (nøytrale / antall_kunder) * 100

# Supportavdelingens NPS
nps = prosent_positive - prosent_negative

# Skriv ut resultatet
print("\n- NPS Beregning -")
print(f"Totalt antall kunder: {antall_kunder}")
print(f"Positive: {positive} ({prosent_positive:.1f}%)")
print(f"Negative: {negative} ({prosent_negative:.1f}%)")
print(f"Nøytrale: {nøytrale} ({prosent_nøytrale:.1f}%)")
print(f"Net Promoter Score (NPS): {nps:.1f}")
