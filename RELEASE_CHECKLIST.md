# Progress Studio V2 Release Checklist

- [x] MS-1 Application Skeleton
- [x] MS-2 Import and Schedule
- [x] MS-3 Timescale and Excel Infrastructure
- [x] MS-4 BOQ and Amount
- [x] MS-5 Progress and Distribution
- [x] MS-6 OKD and Legacy Removal
- [x] MS-7 Regression and Release

## Release command

```bash
python -m unittest discover -s tests -v
python main.py --help
```

## Release identity

- Product: Progress Studio
- Version: 2.0.1
- Primary worksheet: `main`
- Entry point: `python main.py`
- Legacy script dependency: none
