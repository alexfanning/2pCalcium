# 2pCalcium

Python analysis code for in vivo two-photon calcium imaging of individual climbing fiber (CF) axons in awake, head-fixed mice during sensory stimulation and delay eyeblink conditioning.

CFs were labeled by injecting the inferior olive with Cal-520 dextran (calcium indicator, green) and tetramethylrhodamine (TMR) dextran (calcium-insensitive reference, red), and imaged through a cranial window over cerebellar lobules V/VI. The code detects calcium events in single CFs, classifies them as stimulus-evoked or spontaneous, and quantifies event amplitude and frequency. This work was done during my Ph.D. in the Nishiyama Lab at the University of Texas at Austin.

## Pipeline (`ca_data_analysis3_5GR.py`)

Each run processes one CF, one block of 10 trials at a time (6 blocks = 60 trials per session).

1. **Input.** Two text files per CF with mean fluorescence per frame: green (Cal-520) and red (TMR). Traces come from ROIs on motion-corrected movies.
2. **Baseline.** A polynomial baseline (F0) is fit to the green trace for each block with `peakutils.baseline`, so slow drift and bleaching are removed.
3. **Red smoothing.** The red trace is smoothed with a 5-frame moving average to limit green bleed-through into the red channel and reduce noise added by ratioing.
4. **Ratio signal.** ΔG/R = (G − F0) / R_smoothed. ΔF/F0 is also computed for comparison.
5. **Noise estimate.** The standard deviation (SD) of ΔG/R for the block. All detection thresholds are expressed in multiples of this SD.
6. **Event detection.**
   - Candidate peaks are found with `peakutils.indexes`; a peak is shifted one frame later if the next frame is higher.
   - Size filter: peaks smaller than `thres_eli` × SD above the preceding trough are removed (default 1.5).
   - Isolation filter: a peak is kept for amplitude measurement only if its preceding trough is within `thres_iso` × SD of the ΔG/R baseline (default 1), so amplitudes are not inflated by the decay of an earlier event.
7. **Amplitude.** Peak height is measured from the trough 1–2 frames before the peak.
8. **Single-spike reference.** Isolated events smaller than `thres_sin` × SD (default 3) are treated as presumed single-action-potential events.
9. **Classification by peak time.**
   - CS (light) window: 50–200 ms after CS onset.
   - US (air puff) window: 25–150 ms after US onset.
   - Only the first event in each window counts as evoked; later events in the same window are reclassified as spontaneous.
   - On probe trials, events in the CS window of US-only trials and in the US window of CS-only trials are reclassified as spontaneous.
   - All other events are spontaneous.
10. **Normalization.** Each event's amplitude is normalized two ways: to the CF's mean spontaneous amplitude, and to its mean presumed single-spike amplitude.
11. **Outputs.** Diagnostic plots for each block (raw green with baseline, raw and smoothed red, ΔF/F0, ΔG/R with all and isolated peaks marked) and an Excel workbook per CF containing:
    - event counts and event frequency (Hz)
    - per-class amplitudes (CS, US, spontaneous), split into CS-alone vs CS in paired trials and US-alone vs US in paired trials
    - mean, SEM, and SD under both normalizations

### Key parameters

| Parameter | Meaning |
| --- | --- |
| `nb` | Block of 10 trials to analyze (1–6) |
| `nf` | Frames per trial |
| `frame_rate` | Frame period in ms |
| `us_onset` | US onset time within the trial (ms) |
| `cs1`, `us1`, `interval` | Trial numbers of the first CS-only and US-only probe trials, and the spacing between probes |
| `thres_eli`, `thres_iso`, `thres_sin` | Detection thresholds, in multiples of the ΔG/R SD |
| `analysis_complete` | Set to 1 after all blocks are processed to compile the summary sheets |

### Usage

1. Place the green and red trace files (for example `cf1green.txt`, `cf1red.txt`) in the working directory.
2. Set the parameters at the top of the script.
3. Run once per block (`nb = 1` to `6`). Results accumulate in `<filename>.xlsx`.
4. Set `analysis_complete = 1` and run once more to compile the sorted and summary sheets.

## Other scripts

- `ca_data_analysis4_2GR_desensitization.py`: [confirm] variant of the pipeline for the repeated-stimulation experiments, in which stimuli were presented daily across sessions.
- `decodingTime.py`: [describe]
- `sorting_RPMdata2.py`: [describe; for example, organizing running-wheel data for alignment with imaging]

## Requirements

- Python 3
- NumPy
- SciPy
- Matplotlib
- PeakUtils
- openpyxl
- more-itertools

## Related work

Fanning A. Ph.D. dissertation, University of Texas at Austin, 2019.
