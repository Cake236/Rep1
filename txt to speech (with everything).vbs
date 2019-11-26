'*/ To explore Txt to Speech VB Libaray for app embedding use/*'
voic = inputbox("Voice (0 = male and 1 = female)")
stroText = inputbox("Volume (Num.)")
Text = inputbox("Speed (Num.)")
strText = inputbox("Text to Speech (Tex. and/or Num.)")
Set VObj = CreateObject("SAPI.SpVoice")
	with VObj
		Set .voice = .getvoices.item(voic)
		.Volume =stroText
		.Rate = Text
		.Speak strText
	end with



'*/ Hello, now help me to control within audible range for both speed and volume using conditions */'
