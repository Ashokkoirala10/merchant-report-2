from django.db import models


class FonepayMerchantCBS(models.Model):
    merchant_id = models.CharField(max_length=50, unique=True)
    merchant_name = models.CharField(max_length=200, blank=True)
    province = models.CharField(max_length=100, blank=True)
    district = models.CharField(max_length=100, blank=True)
    municipality = models.CharField(max_length=100, blank=True)
    address1 = models.CharField(max_length=300, blank=True)
    address3 = models.CharField(max_length=300, blank=True)
    gender = models.CharField(max_length=20, blank=True)

    class Meta:
        verbose_name = "Fonepay Merchant CBS"


class NepalpayMerchantCBS(models.Model):
    merchant_code = models.CharField(max_length=50)          # removed unique=True
    merchant_account = models.CharField(max_length=50, blank=True, null=True)  # ← NEW FIELD
    
    merchant_name = models.CharField(max_length=200, blank=True)
    province = models.CharField(max_length=100, blank=True)
    district = models.CharField(max_length=100, blank=True)
    municipality = models.CharField(max_length=100, blank=True)
    address1 = models.CharField(max_length=300, blank=True)
    address3 = models.CharField(max_length=300, blank=True)
    gender = models.CharField(max_length=20, blank=True)

    class Meta:
        verbose_name = "Nepalpay Merchant CBS"
        unique_together = [['merchant_code', 'merchant_account']]   # ← This is important

    def __str__(self):
        return f"{self.merchant_code} - {self.merchant_account or 'N/A'}"


class UploadSession(models.Model):
    STATUS_CHOICES = [
        ('uploaded', 'Uploaded'),
        ('processed', 'Processed'),
        ('reviewed', 'Reviewed'),
        ('final', 'Final Report Generated'),
    ]
    created_at = models.DateTimeField(auto_now_add=True)
    month_name = models.CharField(max_length=50, default='Ashwin')
    fonepay_file = models.FileField(upload_to='uploads/', null=True, blank=True)
    nepalpay_file = models.FileField(upload_to='uploads/', null=True, blank=True)
    fonepay_processed = models.FileField(upload_to='processed/', null=True, blank=True)
    nepalpay_processed = models.FileField(upload_to='processed/', null=True, blank=True)
    final_report = models.FileField(upload_to='reports/', null=True, blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='uploaded')
    fonepay_row_count = models.IntegerField(default=0)
    nepalpay_row_count = models.IntegerField(default=0)
