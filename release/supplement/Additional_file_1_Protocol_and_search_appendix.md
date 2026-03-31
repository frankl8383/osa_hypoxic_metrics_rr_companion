# Additional file 1. Protocol and search appendix

This appendix reports the final protocol rules, the exact executed search strings, the final database yields, the eligibility-assessment accounting, and the operational synthesis rules used for the submitted review package. Development-stage logs, pilot counts, local file paths, and intermediate parsing notes are intentionally omitted from this submission version.

## Review question and registration status

- Working review title: `Beyond-AHI hypoxic metrics and hard cardiovascular and mortality outcomes in adult OSA-related cohorts: a systematic review with limited quantitative synthesis`
- Reporting framework: `PRISMA 2020`
- Additional reporting cross-check: `MOOSE-relevant items for observational meta-analyses`
- Registration: protocol outline prepared in a `PROSPERO`-compatible format but not formally registered
- Protocol freeze date: `2026-03-23` before full-text eligibility assessment and before the final limited quantitative synthesis
- Core review question: in adult OSA-related cohorts, which beyond-AHI hypoxic metrics show the strongest and most defensible associations with hard cardiovascular or mortality outcomes?

## Final eligibility framework

Eligible reports were original adult human studies with at least one prespecified beyond-AHI hypoxic metric (`HB`, `SASHB`, `T90`, `ODI`, minimum or nadir oxygen saturation) and at least one hard clinical outcome (all-cause mortality, cardiovascular mortality, incident or adjudicated cardiovascular events, incident heart failure, incident atrial fibrillation, or incident stroke). For interpretation, cohorts were grouped into three classes:

- `clinical OSA/referral cohorts`
- `community cohorts with OSA-related physiology`
- `specialized cardiovascular/surgical cohorts`

Pediatric studies, animal studies, diagnostic-only reports, surrogate-only outcome studies, central-sleep-apnea-focused cohorts without separable OSA results, and non-original or non-extractable reports were excluded from the primary quantitative pool.

## Final synthesis rules

- metric families were kept separate
- categorical contrasts were not pooled with continuous-scale analyses
- materially different composite outcomes were not pooled by default
- adjusted hazard ratios were prioritized
- when multiple estimates from the same cohort family addressed the same metric-outcome cell, default estimate selection prioritized adjusted hazard ratios, prespecified core metric families over comparator constructs, primary published models over alternate incremental models, and the cleanest shared scale before alternate subgroup or threshold estimates
- alternate scales or models from the same cohort family were retained as overlap-sensitive sensitivity evidence rather than merged into the default pooled cell
- random-effects meta-analysis with restricted maximum likelihood was used only for directly comparable cells
- post hoc scale harmonization was labeled exploratory
- because every poolable cell contained fewer than 10 studies, funnel plots and small-study-effects testing were not performed

## Default estimate-selection algorithm

| Priority step | Operational rule |
| --- | --- |
| 1 | Prefer adjusted hazard ratios over unadjusted, logistic, or purely descriptive estimates |
| 2 | Prefer prespecified core metric families over comparator constructs from the same cohort family |
| 3 | Prefer the primary published model over alternate incremental or subtype-adjusted models |
| 4 | Prefer the cleanest shared exposure scale before subgroup-only, threshold-only, or overlap-sensitive estimates |
| 5 | Retain alternate estimates from the same cohort family as sensitivity/comparator evidence rather than merge them into the primary pooled analyses |

## Protocol deviations and late analytic clarifications

- The interpretive framing was narrowed from `adult OSA` to `adult OSA-related cohorts` after full eligibility assessment because several retained community cohorts were not restricted to clinic-defined OSA at entry; the prespecified metric and outcome families were not changed.
- The atrial-fibrillation T90 harmonization was introduced after protocol drafting and is retained only as an exploratory sensitivity analysis.
- Metric-specific PubMed side searches were used as a supplementary recall audit after closure of the main PubMed corpus; they did not create a new primary pooled cell.

## Final search yields and screening position

- database records identified before deduplication: `1636`
- duplicate records removed before screening: `777`
- PubMed main query identified: `513` records
- PubMed records parsed into the screening corpus after export cleaning: `510` records
- PubMed IDs not parsed into the screening corpus after export cleaning: `3` records
- The 3 PubMed IDs not retained in the parsed screening corpus were removed during export cleaning because their saved export blocks did not normalize into complete screening records.
- Web of Science Core Collection exported after document-type filtering: `588` records
- Embase exported after source/publication-type filtering: `535` records
- deduplicated title/abstract screening corpus across the three planned databases: `859` records
- reports sought for retrieval after title/abstract screening: `62`
- reports not retrieved: `0`
- full-text reports assessed for eligibility against the prespecified protocol rules: `62`
- unique articles contributing to the historical executed quantitative evidence set: `24`

Reports not retained in the historical executed quantitative evidence set after full-text assessment fell mainly into three broad groups, while a later-rescued subset was retained into the post-freeze supplement:

- later re-adjudicated into the post-freeze supplement: `7`
- narrative-only or nonextractable reports: `11`
- protocol-scope or intervention-effect-modifier exclusions: `11`
- specialized/context/noncanonical comparator reports that did not yield a prespecified retained quantitative estimate: `13`

## Re-review audit summary

| Audit object | Re-reviewed set | Review structure | Final effect on the submitted core |
| --- | --- | --- | --- |
| Included articles | `31` articles in the final submission dataset (`24` historical + `7` post-freeze upgrade articles) | The second author re-reviewed the retained studies before final submission export, with the post-freeze supplement adjudicated under the same extraction rules | The updated submission dataset was expanded without adding a new primary pooled cell |
| Full-text non-retained reports | `38/38` reports | The second author re-reviewed all non-retained full-text decisions; disagreements were adjudicated by the corresponding author | Final non-retained classes were retained and summarized as final-state exclusion categories |
| Primary pooled inputs | `4/4` pooled cells (`8` cohort-specific estimates) | All effect estimates contributing to the four primary pooled cells were re-checked before export | The four-cell primary pooled structure was retained without change |
| Non-pooled retained estimates | `46/46` retained non-pooled estimates | The second author re-reviewed all retained sensitivity/comparator and narrative-supporting estimates against the extraction worksheet and article PDFs or abstracts | Estimate-level retention outside the primary pooled analyses was preserved while the post-freeze upgrade thickened the T90/TST90 and specialized-evidence layers |

## Post-freeze evidence-upgrade supplement

After the historical executed package was frozen, we performed a targeted post-freeze evidence-upgrade supplement focused on high-value full texts and open-access anchors identified during manuscript finalization. A final strict-review re-adjudication also rescued one previously screened dual-cohort T90 mortality paper into the updated dataset. This supplement is now integrated directly into the final-state Figure 1 accounting and was adjudicated using the same protocol-concordant extraction and retention rules.

- targeted studies/open-access anchors reviewed: `8`
- retained into the updated submission dataset: `7` studies contributing `9` cohort-level estimates
- contextual-only specialized paper reviewed but not included in the quantitative evidence set: `1` (`Pinilla 2023`, PMID `37734857`)
- updated final submission dataset: `31` unique articles, `54` cohort-level estimates, `38` primary estimates, and `16` sensitivity/comparator estimates
- Figure 1 now ends in the final-state dataset while preserving the executed-package subcount: `24` unique articles and `45` cohort-level estimates from the three-database package plus `7` added studies from the integrated supplement
- effect on the four primary pooled cells: `no new pooled cell was added and the four-cell primary pooled structure remained unchanged`

## Final anchor-centered citation-chasing completeness pass

On `2026-03-26`, we completed a targeted anchor-centered citation-chasing completeness pass around the main HB/SASHB mortality or heart-failure anchors, the T90/ODI atrial-fibrillation and mortality anchors, and all studies carried into the post-freeze evidence-upgrade supplement. Candidate follow-on papers were checked against primary-source `PubMed`, `PMC`, or journal DOI records under the same protocol-concordant retention rules.

- anchor set interrogated: `Azarbarzin 2019/2020`, `Labarca 2023`, `Blanchard 2021`, `Baumert 2020`, `Oldenburg 2016`, `Heinzinger 2023`, `Kendzerska 2018`, `Trzepizur 2022`, `Hui 2024`, `Vichova 2025`, `Mazzotti 2025`, plus the retained post-freeze upgrade studies
- strongest rescued article already integrated: `Henríquez-Beltrán 2024` (PMID `37656346`)
- final result of this pass: `no additional protocol-concordant retained study or retained cohort-level estimate beyond the integrated 31-article / 54-estimate updated submission dataset`
- effect on the four primary pooled cells: `no new independent publication-level replication was identified for HB -> cardiovascular mortality, HB -> all-cause mortality, or SASHB -> incident heart failure`
- highest-value screened but non-retained candidates confirmed during this pass:
  - `Xu 2026` (PMID `41794120`): composite high-CVD-risk OSA phenotype based on high HB or high ΔHR rather than a separable prespecified metric-family estimate
  - `Zheng 2025` (PMID `41478496`): specialized ACS pooled cohort with inverse or U-shaped `TSA90` associations rather than a general OSA prognostic anchor
  - `Yan 2024` (PMID `37772691`): community `SpO2_TOTAL` construct rather than a prespecified event-based hypoxic metric family
  - `Parekh 2023` (PMID `37698405`): ventilatory burden, a novel but noncanonical construct outside the prespecified exposure families
- implication: `after the three-database package, the post-freeze upgrade supplement, and the final citation-chasing pass, the main remaining ceiling is publication-level replication scarcity rather than unresolved search incompleteness`

## Supplementary PubMed side-search audit

Metric-specific PubMed side searches were executed as a supplementary recall audit after closure of the main PubMed corpus, not as an independent fourth database stream.

- unique side-search records versus the main PubMed query: `283`
- high-priority audit records: `2`
- medium-priority audit records: `4`
- low-priority audit records: `47`
- likely exclusions on first-pass triage: `230`
- final impact on retained articles or retained cohort-level estimates: `no article or retained estimate in the final quantitative evidence set depended exclusively on the side-search stream`
- effect on the primary pooled analyses: `no new primary pooled cell added`


## Full-text reports not retained after eligibility assessment

| Report | PMID | Context | Final disposition | Main reason for non-retention |
| --- | --- | --- | --- | --- |
| Pencić 2007 | 17633312 | unclear | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Jung 2010 | 20507958 | maintenance hemodialysis | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Masuda 2011 | 21220756 | maintenance hemodialysis pulse-oximetry cohort | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Ensrud 2012 | 22705247 | older community-dwelling men cohort | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Gami 2013 | 23770166 | Minnesota PSG referral | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Aronson 2014 | 24523943 | acute myocardial infarction | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Stone KL 2016 | 26943468 | MrOS Sleep Study | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Xie 2016 | 27464791 | recent myocardial infarction cohort | Later retained in post-freeze supplement | Historical executed package remained frozen; the article was later re-adjudicated and retained in the updated submission dataset. |
| Kendzerska 2016 | 27690206 | suspected OSA obesity-hypoxaemia note | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Yatsu 2018 | 29605831 | post-PCI pulse-oximetry | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Kendzerska 2019 | 30372124 | overlap syndrome | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Huang 2020 | 31967668 | Fuwai ADHF | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Jhamb M 2020 | 31969341 | CKD/ESKD | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Linz 2020 | 32679239 | SAVE trial | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Blanchard 2021 (letter note) | 33214210 | letter note | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Labarca 2021 (SantOSA; PMID 33394326) | 33394326 | SantOSA | Later retained in post-freeze supplement | Historical executed package remained frozen; the article was later re-adjudicated and retained in the updated submission dataset. |
| Cui H 2021 | 34226030 | obstructive HCM septal-myectomy cohort | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Mehra R 2022 | 34797743 | PMID 34797743 | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Huhtakangas JK 2022 | 35679775 | acute ischemic stroke | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Zapater 2022 | 35833104 | ISAACC | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Cardoso 2023 | 36690808 | resistant hypertension | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Mochida 2023 | 36769506 | maintenance hemodialysis | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Varol 2024 | 37422579 | diagnosed OSA PSG | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Esmaeili N 2023 | 37531573 | SHHS | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Henríquez-Beltrán 2024 | 37656346 | SHHS and SantOSA | Later retained in post-freeze supplement | Historical executed package remained frozen; the article was later re-adjudicated and retained in the updated submission dataset. |
| Huang B 2023 | 37724626 | hospitalized HF | Specialized/context/noncanonical comparator report | Did not yield a prespecified retained quantitative estimate in the historical executed evidence set. |
| Pinilla L 2023 | 37734857 | ISAACC trial | Protocol exclusion | Intervention or effect-modifier analysis rather than prognostic cohort evidence. |
| Yan B 2024 | 37772691 | SHHS | Protocol exclusion | Community-based nocturnal saturation construct lay outside the prespecified OSA-related metric framework. |
| Azakli 2024 (letter note) | 38097477 | letter note | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Nyuta E 2024 | 38749745 | AF ablation | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Briancon-Marjollet 2024 | 38978551 | non-obese OSA | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Azarbarzin A 2025 | 40794640 | RICCADSA / ISAACC / SAVE pooled analysis | Protocol exclusion | Intervention or effect-modifier analysis rather than prognostic cohort evidence. |
| Potratz M 2025 | 41306066 | heart transplant waitlist | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |
| Zheng W 2025 | 41478496 | two ACS prospectives | Specialized/context/noncanonical comparator report | Specialized ACS secondary-prevention cohort with inverse or U-shaped TSA90 associations; not retained as a general OSA prognostic anchor. |
| Xu 2026 | 41794120 | MESA | Specialized/context/noncanonical comparator report | Composite high-risk OSA phenotype combined HB or ΔHR rather than reporting a separable prespecified metric-family estimate. |
| Labarca 2021 (SantOSA; PMID 32232718) | 32232718 | SantOSA | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Sutherland 2022 | 35896039 | Sleep Heart Health Study | Narrative-only or nonextractable report | No extractable protocol-concordant hard-outcome estimate for the historical executed quantitative evidence set. |
| Lowery 2023 | 37968017 | PVDOMICS group 1 PAH | Protocol exclusion | Outside the prespecified OSA-related prognostic scope. |

Complete report-level log for all 38 full-text reports that were reviewed but not retained in the historical executed quantitative evidence set.


## Executed PubMed main query

```text
("Sleep Apnea, Obstructive"[Mesh] OR "Sleep Apnea Syndromes"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab] OR OSA[tiab])
AND
("hypoxic burden"[tiab] OR "oxygen desaturation index"[tiab] OR ODI[tiab] OR T90[tiab] OR "sleep hypoxemia"[tiab] OR "sleep hypoxaemia"[tiab] OR "nocturnal hypoxemia"[tiab] OR "nocturnal hypoxaemia"[tiab] OR "minimum oxygen saturation"[tiab] OR "nadir oxygen saturation"[tiab] OR "lowest oxygen saturation"[tiab] OR "nadir SpO2"[tiab])
AND
(mortality[tiab] OR "cardiovascular mortality"[tiab] OR "all-cause mortality"[tiab] OR "cardiovascular event*"[tiab] OR MACE[tiab] OR "major adverse cardiovascular"[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR prognosis[tiab] OR incident[tiab])
NOT
(child*[tiab] OR pediatric*[tiab] OR paediatric*[tiab])
```


## Executed supplementary PubMed side-search queries

```text
HB:
("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
AND ("hypoxic burden"[tiab])
AND (mortality[tiab] OR "cardiovascular mortality"[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

T90 / sleep hypoxemia:
("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
AND (T90[tiab] OR "sleep hypoxemia"[tiab] OR "sleep hypoxaemia"[tiab] OR "time below 90"[tiab] OR "minimum oxygen saturation below 90"[tiab])
AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

ODI:
("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
AND ("oxygen desaturation index"[tiab] OR ODI[tiab])
AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])

Nadir / minimum oxygen saturation:
("Sleep Apnea, Obstructive"[Mesh] OR "obstructive sleep apnea"[tiab] OR "obstructive sleep apnoea"[tiab] OR "sleep-disordered breathing"[tiab])
AND ("minimum oxygen saturation"[tiab] OR "nadir oxygen saturation"[tiab] OR "lowest oxygen saturation"[tiab] OR "nadir SpO2"[tiab])
AND (mortality[tiab] OR "heart failure"[tiab] OR "atrial fibrillation"[tiab] OR stroke[tiab] OR cardiovascular[tiab])
```


## Executed Web of Science Core Collection query

```text
TS=((("obstructive sleep apnea" OR "obstructive sleep apnoea" OR "sleep-disordered breathing" OR "sleep disordered breathing" OR OSA OR OSAHS) AND ("hypoxic burden" OR "sleep apnea-specific hypoxic burden" OR "sleep apnoea-specific hypoxic burden" OR "oxygen desaturation index" OR ODI OR T90 OR "sleep hypoxemia" OR "sleep hypoxaemia" OR "nocturnal hypoxemia" OR "nocturnal hypoxaemia" OR "minimum oxygen saturation" OR "nadir oxygen saturation" OR "lowest oxygen saturation" OR "nadir SpO2" OR "mean oxygen saturation") AND (mortality OR "all-cause mortality" OR "cardiovascular mortality" OR "cardiovascular event*" OR MACE OR MACCE OR "major adverse cardiovascular" OR "heart failure" OR "atrial fibrillation" OR stroke OR prognosis OR incident)) NOT (child* OR pediatric* OR paediatric*))
```

Post-retrieval filters:

- database: `Web of Science Core Collection`
- document types: `Article`, `Early Access`


## Executed Embase query logic

```text
#1 ('obstructive sleep apnea'/exp OR 'obstructive sleep apnea':ti,ab,kw OR 'obstructive sleep apnoea':ti,ab,kw OR 'sleep disordered breathing':ti,ab,kw OR 'sleep-disordered breathing':ti,ab,kw OR osa:ti,ab,kw OR osahs:ti,ab,kw)

#2 ('hypoxic burden'/exp OR 'hypoxic burden':ti,ab,kw OR 'sleep apnea specific hypoxic burden':ti,ab,kw OR 'sleep apnoea specific hypoxic burden':ti,ab,kw OR 'oxygen desaturation index'/exp OR 'oxygen desaturation index':ti,ab,kw OR odi:ti,ab,kw OR t90:ti,ab,kw OR 'sleep hypoxemia':ti,ab,kw OR 'sleep hypoxaemia':ti,ab,kw OR 'nocturnal hypoxemia':ti,ab,kw OR 'nocturnal hypoxaemia':ti,ab,kw OR 'minimum oxygen saturation':ti,ab,kw OR 'nadir oxygen saturation':ti,ab,kw OR 'lowest oxygen saturation':ti,ab,kw OR 'nadir spo2':ti,ab,kw OR 'mean oxygen saturation':ti,ab,kw)

#3 ('mortality'/exp OR mortality:ti,ab,kw OR 'all cause mortality':ti,ab,kw OR 'cardiovascular mortality':ti,ab,kw OR 'cardiovascular event*':ti,ab,kw OR mace:ti,ab,kw OR macce:ti,ab,kw OR 'major adverse cardiovascular':ti,ab,kw OR 'heart failure'/exp OR 'heart failure':ti,ab,kw OR 'atrial fibrillation'/exp OR 'atrial fibrillation':ti,ab,kw OR stroke/exp OR stroke:ti,ab,kw OR prognosis/exp OR prognosis:ti,ab,kw OR incident:ti,ab,kw)

#4 (child/exp OR adolescent/exp OR pediatric*:ti,ab,kw OR paediatric*:ti,ab,kw OR child*:ti,ab,kw)

Final logic: #1 AND #2 AND #3 NOT #4
```

Post-retrieval filters:

- publication types: `Article`, `Article in Press`
- source limit: `[embase]/lim`


## Deduplication and evidence-layer notes

- Deduplication was performed sequentially against the closed corpus using PMID, DOI, and normalized title matching, followed by manual review of residual ambiguous records.
- The final quantitative evidence set was built at the cohort-row level rather than the article level.
- Alternate models or scales from the same cohort family were retained as overlap-sensitive sensitivity evidence rather than merged into the primary pooled analyses.
- Narrative-only and noncanonical comparator records were tracked separately from prespecified metric-family estimates.

