# ParsingWeatherDataWithPowerPoint_Example
파워포인트의 매크로 기능을 이용하여 전국 날씨를 네이버으로 부터 파싱하는 예제 입니다.

![미리보기](https://user-images.githubusercontent.com/40740128/68759981-253d1100-0654-11ea-82cd-1057d37e4ec8.gif)

[[더 많은 이미지를 보시려면 클릭]](https://blog.naver.com/sungbin_dev/221706873436)

## SourceCode
``` VBA
Function GetHTML(URL As String) As String
    Dim Html As String
    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", URL, False
        .Send
        GetHTML = .ResponseText
    End With
End Function
Function ReplaceRegex(Text As String, Regex As String, NewText As String) As String
    Dim RegexObject As Object
    Set RegexObject = CreateObject("vbscript.regexp")

    With RegexObject
        .Pattern = Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With

    ReplaceRegex = RegexObject.Replace(Text, NewText)
End Function
Function SplitText(Text As String, Regex As String, Index As Integer) As String
    Dim Data() As String
    Data = Split(Text, Regex)
    SplitText = Data(Index)
End Function
Function JoinEnter(Text As String) As String
    Text = Replace(Text, "백령", "[백령]")
    Dim AreaList As Variant
    AreaList = Array("서울", "춘천", "강릉", "대전", "청주", "전주", "대구", "광주", "부산", "제주", "울릉/독도", "안동", "목포", "여수", "울산", "수원")
    For i = 0 To 15
        Dim Area As String
        Area = AreaList(i)
        Text = Replace(Text, Area, "[" + Area + "]")
        Text = Replace(Text, "[" + Area + "]", vbCrLf + "[" + Area + "]")
    Next i
    JoinEnter = Text
End Function
Private Sub LoadWeatherView_Click()
    WeatherView.Caption = "전국날씨 불러오는중..."
    Dim Value As String
    Value = GetHTML("https://m.search.naver.com/search.naver?query=전국날씨")
    Value = SplitText(Value, "전국날씨</strong>", 1)
    Value = SplitText(Value, "<div class=""t_notice"">", 0)
    Value = ReplaceRegex(Value, "<!*[^<>]*>", "")
    Value = Trim(Value)
    Value = JoinEnter(Value)
    Value = Replace(Value, "  ", "")
    Value = Replace(Value, "도씨", "℃")
    Value = SplitText(Value, "관련날씨뉴스", 0)
    Value = Replace(Value, "단위 ℃특보", "")
    Value = Replace(Value, "(", vbCrLf + "기상특보 (")
    Value = Replace(Value, ") ", ")" + vbCrLf)
    Value = Replace(Value, "기상특보", vbCrLf + "기상특보", 1, 1)
    Value = Replace(Value, "기준기상청", "기상청 발표 기준}")
    Dim NowDate As String
    NowData = Format(Date, "mm.dd")
    Value = Replace(Value, NowData, vbCrLf + vbCrLf + "{" + NowData)
    WeatherView.Caption = Value
End Sub
```

## Example PPTM File Download - [[CLICK]](https://github.com/sungbin5304/ParsingWeatherDataWithPowerPoint_Example/raw/master/%ED%8C%8C%EC%9B%8C%ED%8F%AC%EC%9D%B8%ED%8A%B8%EB%A1%9C%20%EB%82%A0%EC%94%A8%20%ED%8C%8C%EC%8B%B1%ED%95%98%EA%B8%B0.pptm)
