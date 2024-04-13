from django.urls import path
from .views import home,AboutUs,ContactUs,UserFeedback,PdfToWordConverter,DocxToPdfConverter,ImgToPdfConverter,PdfToImgConverter,\
    PdfCompresser,MergePdfs,PptToPdfConverter,PdfEncrypter,login,register,introduction
urlpatterns = [
    path('', home, name='ptw'),
    path('log/',login,name='login'),
    path('reg/',register,name='regi'),
    path('intro1',introduction, name='intro'),
    path('AboutUs/', AboutUs, name='About'),
    path('ContactUs/', ContactUs, name='contact'),
    path('FeedBack/', UserFeedback, name='feedback'),
    path('PDF_TO_WORD_CONVERTED/',PdfToWordConverter,name='ptwcd'),
    path('WORD_TO_PDF_CONVERTED/',DocxToPdfConverter,name='wtpcd'),
    path('Image_To_PDF_CONVERTED/',ImgToPdfConverter,name='itpcd'),
    path('PDF_TO_IMAGE_CONVERTED/',PdfToImgConverter,name='pticd'),
    path('PDF_COMPRESSED/',PdfCompresser,name='pc'),
    path('filepicker/',MergePdfs,name='filepicker'),
    path('PPT_TO_PDF_CONVERTED/',PptToPdfConverter,name='ptppd'),
    path('PDF_ENCRYPTED/',PdfEncrypter,name='pe'),

]