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



'hello
