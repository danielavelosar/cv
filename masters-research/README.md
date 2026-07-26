# Recherche masters & financement — rentrée 2027

Recherche orchestrée exécutée le **2026-07-26** selon le métaprompt « 4 phases × 3 langues » : 12 passes de recherche (FR/EN/ES), 4 agents de comparaison, un deck final et une étape d'auto-réfutation.

## Livrables

| Fichier | Contenu |
|---|---|
| **`masters-2027-deck.pptx`** | Le deck final (194 diapositives) : légende des badges, synthèse exécutive, 10 blocs domaines, transversal (EMJM, Chevening, financement colombien, rétroplanning, registres 🔴/⚫), annexes (fiches programmes + bibliographies) |
| `phase-1-dossier.md` | Dossier consolidé — HCI · Visualisation sci./médicale · Serious games |
| `phase-2-dossier.md` | Dossier consolidé — IA santé · Neuro computationnelle/BCI · Bioinformatique |
| `phase-3-dossier.md` | Dossier consolidé — Biomécanique/mouvement · Animation 3D TD · VFX/R&D |
| `phase-4-dossier.md` | Dossier consolidé — Économie quantitative + synthèses transversales (tableau maître du financement, rétroplanning) |
| `contre-argumentaire.md` | L'étape que la plupart des gens sautent : arguments CONTRE les 3 recommandations + fragilités du deck |
| `sources.md` | 633 URLs consultées, datées, par passe |
| `raw/phase-{1..4}-{FR,EN,ES}.md` | Les 12 rapports bruts par langue |
| `deck/` | Générateur du deck (`build-deck.js` + `data/phase-{1..4}.json`) |

## Avertissement méthodologique

L'accès direct aux pages officielles était largement bloqué (403) pendant la recherche : la plupart des faits proviennent de résumés de moteur de recherche citant des URL officielles, et sont donc badgés 🟡 (« à confirmer »). Presque aucune date du cycle 2027 n'était publiée en juillet 2026. **Chaque donnée décisive doit être relue à l'URL source avant toute candidature.** Les registres 🔴 (contradictions) et ⚫ (non trouvés) figurent dans chaque dossier et dans le deck.

## Regénérer le deck

```bash
cd masters-research/deck
npm install pptxgenjs --no-save
node build-deck.js   # produit ../masters-2027-deck.pptx
```
