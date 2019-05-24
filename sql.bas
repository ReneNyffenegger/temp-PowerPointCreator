option explicit


dim pres as powerPoint.presentation

sub addSlide(byVal text as string) ' {

    dim slid   as powerPoint.slide
    dim shText as powerPoint.shape

    text = replace(text, chr(13), chr(13) & "SQL> ")

    set slid = pres.slides.add(pres.slides.count + 1, ppLayoutBlank)
    slid.followMasterBackground = false
    slid.background.fill.solid
    slid.background.fill.foreColor.rgb = rgb(15, 175, 20)

    set shText = slid.shapes.addTextBox(msoTextOrientationHorizontal, 50, 50, 1000, 500)

    shText.textFrame.textRange           =  text
    shText.textFrame.textRange.font.name = "Courier New"
    shText.textFrame.textRange.font.bold =  true
    shText.textFrame.textRange.font.size =    54
    shText.textFrame.textRange.font.color.rgb = rgb(245, 210, 30)

    slid.slideShowTransition.advanceTime   = 0.02 + rnd() / 12
    slid.slideShowTransition.advanceOnTime = true

end sub ' }

sub main() ' {

    set pres = application.presentations.add

    dim sql as string

    sql = "SQL> begin transaction;" & chr(13) & "type, type, type;" & chr(13) & "hack more;" & chr(13) & "delete from customers;" & chr(13) & "rollback;" & chr(13) & ":-)"

    dim i as long

    for i = 1 to len(sql)
        call addSlide(left(sql, i))
    next i

    for i = 1 to 10
        call addSlide(sql)
    next i

  ' pres.slideMaster.background.fill.backColor.rgb = rgb(100, 255, 125)

end sub ' }
