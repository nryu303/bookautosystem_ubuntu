"""enums のユニットテスト"""

from src.common.enums import (
    AvailabilityStatus,
    OrderStatus,
    ParseStatus,
    HoldBucketStatus,
    SupplierCategory,
    UrlMode,
    MailSendStatus,
)


class TestEnums:
    def test_availability_status_values(self):
        assert AvailabilityStatus.AVAILABLE.value == "AVAILABLE"
        assert AvailabilityStatus.UNAVAILABLE.value == "UNAVAILABLE"
        assert AvailabilityStatus.UNKNOWN.value == "UNKNOWN"
        assert AvailabilityStatus.ERROR.value == "ERROR"

    def test_order_status_values(self):
        assert OrderStatus.PENDING.value == "PENDING"
        assert OrderStatus.SELF_STOCK.value == "SELF_STOCK"
        assert OrderStatus.HOLD.value == "HOLD"
        assert OrderStatus.ORDERED.value == "ORDERED"

    def test_enum_string_comparison(self):
        assert AvailabilityStatus.AVAILABLE == "AVAILABLE"
        assert OrderStatus.PENDING == "PENDING"

    def test_parse_status(self):
        assert ParseStatus.PENDING.value == "PENDING"
        assert ParseStatus.DONE.value == "DONE"
        assert ParseStatus.ERROR.value == "ERROR"

    def test_hold_bucket_status(self):
        assert HoldBucketStatus.ACTIVE.value == "ACTIVE"
        assert HoldBucketStatus.THRESHOLD_REACHED.value == "THRESHOLD_REACHED"

    def test_supplier_category(self):
        assert SupplierCategory.EC.value == "ec"
        assert SupplierCategory.STORE.value == "store"

    def test_url_mode(self):
        assert UrlMode.TEMPLATE.value == "template"
        assert UrlMode.CSV_LOOKUP.value == "csv_lookup"
