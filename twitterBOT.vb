Sub oto_beğeni_retweet()  'isim'

'Kodların sonunda yazan rakamlar  Saat+Dakika+Saniye olmaktadır işlemi yapma saniyesidir mesela 5 saniye sonra j tuşuna basacak ondan 5 saniye sonra I '

ActiveWorkbook.FollowHyperlink Address:="Buraya Gidilecek Hashtag Linki"
Application.Wait (Now + TimeValue("00:00:10")) ' Başlat tuşuna bastığınızdan 10 saniye sonra sayfayı açıp işleme başlacaktır.'

For i = 1 To 10 'Burası Beğenilecek tweetleri mesela 10 tane tweet beğenecek siz isterseniz çok yüksek rakamlarda yapabilirsiniz'

Call SendKeys("j", True)
Application.Wait (NowTimeValue("00:00:05")) 'Anlattığım gibi sayfaya gidince burada j tuşuna basacak ve tweeti seçecektir.  '

Call SendKeys("l", True)
Application.Wait (Now + TimeValue("00:00:05"))  'Belirlediğimiz Tweeti beğenecektir '

'Devamını kafanıza göre getirebilirsiniz fakat retweet yaparken tıkladıktan sonra
 Enter tuşuna tıklamanız gerekiyor burada Enter simgesi (~)dir  ALT GR + Ü'

Next
End Sub 'Bitiş'