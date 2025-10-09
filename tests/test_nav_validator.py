from datetime import date, datetime
from decimal import Decimal

from nav_checker.domain.models import NavIdentity, NavRecord
from nav_checker.domain.services import NavValidator


def make_record(instrument: str, nav_value: str) -> NavRecord:
    identity = NavIdentity(
        instrument_id=instrument,
        nav_date=date(2024, 1, 1),
        currency="USD",
        share_class="A",
    )
    return NavRecord(
        identity=identity,
        nav=Decimal(nav_value),
        source="test",
        file_hash="hash",
        as_of=datetime.utcnow(),
    )


def test_no_differences():
    validator = NavValidator()
    inbound = [make_record("ABC", "1.0")]
    authoritative = [make_record("ABC", "1.0")]

    report = validator.compare(inbound, authoritative)

    assert report.summary.wrong_numbers == 0
    assert report.summary.missing_in_inbound == 0
    assert report.summary.unauthorized_in_inbound == 0
    assert not report.has_issues()


def test_wrong_number_detected():
    validator = NavValidator()
    inbound = [make_record("ABC", "1.0")]
    authoritative = [make_record("ABC", "2.0")]

    report = validator.compare(inbound, authoritative)

    assert report.summary.wrong_numbers == 1
    assert report.has_issues()


def test_missing_and_unauthorized():
    validator = NavValidator()
    inbound = [make_record("XYZ", "1.0")]
    authoritative = [make_record("ABC", "1.0")]

    report = validator.compare(inbound, authoritative)

    assert report.summary.missing_in_inbound == 1
    assert report.summary.unauthorized_in_inbound == 1


def test_duplicate_detection():
    validator = NavValidator()
    inbound = [make_record("ABC", "1.0"), make_record("ABC", "1.0")]
    authoritative = [make_record("ABC", "1.0")]

    report = validator.compare(inbound, authoritative)

    assert report.summary.duplicates == 1
    assert report.has_issues()
