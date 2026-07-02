# 2pCalcium

Python scripts for analyzing calcium dynamics from in vivo two-photon imaging of single neurons in awake-behaving mice.

## Overview

This repository contains Python analysis pipelines developed for characterizing graded calcium signal dynamics in brainstem neurons during sensorimotor behavior. The code was used to analyze two-photon calcium imaging data acquired simultaneously with patch-clamp electrophysiology, enabling direct comparison of fluorescence-based and electrical measures of neural activity.

## Contents

- `ca_data_analysis3_5GR.py` — Calcium signal analysis pipeline for 5-group experimental designs; includes signal extraction, baseline normalization, trial averaging, and statistical comparison across conditions
- `ca_data_analysis4_2GR_desensitization.py` — Analysis pipeline for 2-group desensitization experiments; quantifies adaptation of calcium responses across repeated stimulus presentations
- `decodingTime.py` — Temporal decoding analysis; estimates the time course over which calcium signals carry information about behavioral or stimulus variables
- `sorting_RPMdata2.py` — Data sorting and preprocessing for rotary position monitor (RPM) behavioral data; structures raw behavioral data for alignment with neural recordings

## Requirements

- Python 3.7+
- NumPy
- SciPy
- Matplotlib
- Pandas

## Related Publication

**Fanning A**, Shakhawat A, Raymond J. Population calcium responses of Purkinje cells in the oculomotor cerebellum driven by non-visual input. *Journal of Neurophysiology*, 2021;126(4):1391–1402.
