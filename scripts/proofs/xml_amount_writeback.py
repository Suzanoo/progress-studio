"""Engineering proof harness, not a production CLI/Desktop workflow.

First invocation prints an assessment and its review digest. Inspect it, then
repeat with --review-digest DIGEST and --output NEW.xml. A changed input/options
produces a different digest and cannot be approved by the old review.
"""
import argparse
from dataclasses import asdict
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from progress_studio.services.xml_writeback_service import WritebackRequest, XmlWritebackService


def assessment_json(assessment):
    def convert(value):
        if isinstance(value, (Path, Decimal)):
            return str(value)
        raise TypeError(type(value).__name__)
    payload = asdict(assessment)
    payload.update(valid=assessment.valid, replacements=assessment.replacement_count,
                   unchanged=assessment.unchanged_count, preserved_milestones=assessment.preserved_milestone_count)
    return json.dumps(payload, default=convert, ensure_ascii=False, sort_keys=True, indent=2)


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook',type=Path,required=True)
    parser.add_argument('--target-xml',type=Path,required=True)
    field=parser.add_mutually_exclusive_group(required=True)
    field.add_argument('--field')
    field.add_argument('--create-field',action='store_true')
    parser.add_argument('--monetary-confirmed',action='store_true')
    parser.add_argument('--recalculated-and-saved',action='store_true')
    parser.add_argument('--review-digest')
    parser.add_argument('--output',type=Path)
    args=parser.parse_args(argv)
    service=XmlWritebackService()
    assessment=service.assess(WritebackRequest(args.workbook,args.target_xml,args.field,args.create_field,
                                              args.monetary_confirmed,args.recalculated_and_saved))
    serialized=assessment_json(assessment)
    digest=hashlib.sha256(serialized.encode()).hexdigest()
    print(serialized)
    print('Review digest:',digest)
    if not assessment.valid:
        return 2
    if args.output is not None or args.review_digest is not None:
        if not args.output or args.review_digest != digest:
            print('Export blocked: supply a new output path and the exact reviewed digest.',file=sys.stderr)
            return 2
        try:
            result=service.export(assessment,args.output,reviewed=True)
        except (OSError,ValueError) as exc:
            print('Export blocked:',exc,file=sys.stderr)
            return 2
        print('Output:',result.output)
        print('Output SHA256:',result.sha256)
        for warning in result.warnings:
            print('Warning:',warning)
    return 0


if __name__=='__main__':
    raise SystemExit(main())
