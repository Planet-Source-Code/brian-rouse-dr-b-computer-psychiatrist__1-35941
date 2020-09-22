Attribute VB_Name = "Module1"

Public Function Greeting() As String
 Dim Choice As Integer
  Randomize
  Choice = CInt(6 * Rnd + 1)
  Select Case Choice
    Case 1
      Greeting = "Hello, how may I help you?"
    Case 2
      Greeting = "Greetings. What would you like to talk about?"
    Case 3
      Greeting = "Good day. Please tell me your problems."
    Case 4
      Greeting = "What is on your mind today?"
    Case 5
      Greeting = "Please begin when you are ready."
    Case Else
      Greeting = "Hello, what is your question?"
  End Select
End Function

Public Function NewReply(LastReply As String, Question As String) As String
 Dim Choice As Integer
 Dim Location As Integer
 Dim TempReply As String
  Randomize
  Question = UCase(Question)

  Do ' This Do-Until loop keeps Dr.B from repeating himself
    Do
  
  ' Check for "How", "Who", "What", "When", "Why", or "Where" keywords
    If ((InStr(Question, "HOW")) > 0) Or ((InStr(Question, "WHO")) > 0) Or ((InStr(Question, "WHAT")) > 0) Or ((InStr(Question, "WHEN")) > 0) Or ((InStr(Question, "WHERE")) > 0) Or ((InStr(Question, "WHY")) > 0) Then
      Choice = CInt(9 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Why do you ask?"
        Case 2
          NewReply = "Does that question interest you?"
        Case 3
          NewReply = "What answer would please you the most?"
        Case 4
          NewReply = "Are such questions on your mind often?"
        Case 5
          NewReply = "What is it that you really want to know?"
        Case 6
          NewReply = "Why do you ask?"
        Case 7
          NewReply = "Why?"
        Case 8
          NewReply = "What?"
        Case Else
          NewReply = "What do you think?"
      End Select
      Exit Do
    End If
  
  ' Check for "Mother", "Father", "Brother", "Sister", or "Family" keywords
    If ((InStr(Question, "MOTHER")) > 0) Or ((InStr(Question, "BROTHER")) > 0) Or ((InStr(Question, "WHAT")) > 0) Or ((InStr(Question, "WHEN")) > 0) Or ((InStr(Question, "SISTER")) > 0) Or ((InStr(Question, "FAMILY")) > 0) Then
      Choice = CInt(9 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Why do you mention your family?"
        Case 2
          NewReply = "Did you get along with your family?"
        Case 3
          NewReply = "How does your family treat you?"
        Case 4
          NewReply = "Were you ever close to your family?"
        Case 5
          NewReply = "Did you have a happy childhood?"
        Case 6
          NewReply = "Do you like your family members?"
        Case 7
          NewReply = "Did your family ask you to talk to me?"
        Case 8
          NewReply = "What kind of childhood did you have?"
        Case Else
          NewReply = "Do you think about your family often?"
      End Select
      Exit Do
    End If
    
    If ((InStr(Question, "HUSBAND")) > 0) Then
        Choice = CInt(2 * Rnd + 1)
        Select Case Choice
            Case 1
                NewReply = "You need to let God worry about your husband!"
            Case Else
                NewReply = "Honey you are too screwed up in the head right now for a husband"
            End Select
            Exit Do
        End If
        
  If ((InStr(Question, "WIFE")) > 0) Then
        Choice = CInt(2 * Rnd + 1)
        Select Case Choice
            Case 1
                NewReply = "You need to let God worry about your wife!"
            Case Else
                NewReply = "My brother, you are too screwed up in the head right now for a wife"
            End Select
            Exit Do
        End If
  
  ' Check for "Hello" answer
    If ((InStr(Question, "HELLO")) > 0) Or ((InStr(Question, "HI")) > 0) Or ((InStr(Question, "WHATS UP")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "How do you do. How can I help you?"
        Case 2
          NewReply = "Greetings to you too. What shall we talk about today?"
        Case 3
          NewReply = "Bon jour. Shall we get to business now?"
        Case Else
          NewReply = "Please make yourself comfortable and let's begin, shall we?"
      End Select
      Exit Do
    End If
  
  ' Check for "Nothing" answer
    If ((InStr(Question, "NOTHING")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Nothing?"
        Case 2
          NewReply = "I did not ask you what you were worth!"
        Case 3
          NewReply = "Something from nothing leaves nothing!"
        Case Else
          NewReply = "You are a real loser!"
      End Select
      Exit Do
    End If
  
  ' Check for profanity
    If ((InStr(Question, "FUCK")) > 0) Or ((InStr(Question, "SHIT")) > 0) Or ((InStr(Question, "HELL")) > 0) Or ((InStr(Question, "DAMN")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Please don't use four-letter words."
        Case 2
          NewReply = "Profanity is not necessary."
        Case 3
          NewReply = "Do you use such foul language often?"
        Case Else
          NewReply = "You don't have to use profanity to express yourself."
      End Select
      Exit Do
    End If
  
  ' Check for Dream Keyword
    If ((InStr(Question, "DREAM")) > 0) Or ((InStr(Question, "FANTASY")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Wow!"
        Case 2
          NewReply = "You are really in your fantasy world!"
        Case 3
          NewReply = "So your crazy and into witchcraft!"
        Case Else
          NewReply = "You're Pathetic!"
      End Select
      Exit Do
    End If
  
  ' Check for "Name" keyword
    If ((InStr(Question, "NAME")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "As Shakespeare said, 'What's in a name?'"
        Case 2
          NewReply = "Are names important to you?"
        Case 3
          NewReply = "Why do you mention names at all?"
        Case Else
          NewReply = "Names are not important at this time."
      End Select
      Exit Do
    End If ' Check for "yes" answer
    If ((InStr(Question, "YES")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "I knew you would answer that question with a yes."
        Case 2
          NewReply = "You seem to be a positive person am I correct?"
        Case 3
          NewReply = "Is that all you can say is Yes?"
        Case Else
          NewReply = "Have you ever said no to anything?"
      End Select
      Exit Do
    End If
' Check for "KNOW" answer
    If ((InStr(Question, "KNOW")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "What do you know?"
        Case 2
          NewReply = "Don't you know?"
        Case 3
          NewReply = "If you knew so much then you would not be talking to me!"
        Case Else
          NewReply = "Know is short for Knowledge!"
      End Select
      Exit Do
    End If
' Check for "No" answer
    If ((InStr(Question, "NO")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "I knew you would answer that question with a no."
        Case 2
          NewReply = "You seem to be a negative person am I correct?"
        Case 3
          NewReply = "Is that all you can say is no?"
        Case Else
          NewReply = "Have you ever said yes to anything?"
      End Select
      Exit Do
    End If
    ' Check for "money" answer
    If ((InStr(Question, "MONEY")) > 0) Or ((InStr(Question, "CASH")) > 0) Or ((InStr(Question, "FUNDS")) > 0) Or ((InStr(Question, "BILLS")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Money isn't everything!"
        Case 2
          NewReply = "Do you pay your bills on time?"
        Case 3
          NewReply = "How much money to you have?"
        Case Else
          NewReply = "You seem like the broke type!"
      End Select
      Exit Do
    End If

    

' Check for "goodbye" answer
    If ((InStr(Question, "GOODBYE")) > 0) Or ((InStr(Question, "BYE")) > 0) Or ((InStr(Question, "LATER")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Are you sure you want to end your session?"
        Case 2
          NewReply = "I think you are a very disturbed individual, Bye!"
        Case 3
          NewReply = "As they say in Italian, Arrivederci, that means goodbye!"
        Case Else
          NewReply = "See ya and I wouldn't wanna be ya!"
      End Select
      Exit Do
    End If
     
  ' Check for "Thank" keyword
    If ((InStr(Question, "THANK")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "You're welcome."
        Case 2
          NewReply = "No problem."
        Case 3
          NewReply = "I'm always glad to be of service to you."
        Case Else
          NewReply = "Do you feel better now?"
      End Select
      Exit Do
    End If
  
  'Check for "Jesus" answer
    If ((InStr(Question, "JESUS")) > 0) Or ((InStr(Question, "CHURCH")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "DONT MAKE ME START PREACHIN NOW!"
        Case 2
          NewReply = "I love the Lord!"
        Case 3
          NewReply = "Jesus is Lord!"
        Case Else
          NewReply = "I have been raised Pentacostal myself!"
      End Select
      Exit Do
    End If
  
  'Check for "Lord" answer
    If ((InStr(Question, "LORD")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "DONT MAKE ME START PREACHIN NOW!"
        Case 2
          NewReply = "I love the Lord!"
        Case 3
          NewReply = "Jesus is Lord!"
        Case Else
          NewReply = "I have been raised Pentacostal myself!"
      End Select
      Exit Do
    End If
  
  'Check for "MASTER" answer
    If ((InStr(Question, "MASTER")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "DONT MAKE ME START PREACHIN NOW!"
        Case 2
          NewReply = "I love the Lord!"
        Case 3
          NewReply = "Jesus is Lord!"
        Case Else
          NewReply = "I have been raised Pentacostal myself!"
      End Select
      Exit Do
    End If
  'Check for "SAVIOR" answer
    If ((InStr(Question, "SAVIOR")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "DONT MAKE ME START PREACHIN NOW!"
        Case 2
          NewReply = "I love the Lord!"
        Case 3
          NewReply = "Jesus is Lord!"
        Case Else
          NewReply = "I have been raised Pentacostal myself!"
      End Select
      Exit Do
    End If
  
  'Check for "BIBLE" answer
    If ((InStr(Question, "BIBLE")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "DONT MAKE ME START PREACHIN NOW!"
        Case 2
          NewReply = "I love the Lord!"
        Case 3
          NewReply = "Jesus is Lord!"
        Case Else
          NewReply = "I have been raised Pentacostal myself!"
      End Select
      Exit Do
    End If
  
  
  
  ' Check for "Cause" keyword
    If ((InStr(Question, "CAUSE")) > 0) Then
      Choice = CInt(4 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Is that the real reason?"
        Case 2
          NewReply = "Do any other reasons come to mind?"
        Case 3
          NewReply = "Does that reason explain anything else?"
        Case Else
          NewReply = "What other reasons might there be?"
      End Select
      Exit Do
    End If
  
  ' Check for "Sorry" keyword
    If ((InStr(Question, "SORRY")) > 0) Then
      Choice = CInt(5 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Please don't feel the need to apologize."
        Case 2
          NewReply = "Apologies are not necessary."
        Case 3
          NewReply = "How do you feel when you apologize?"
        Case 4
          NewReply = "Don't be so defensive."
        Case Else
          NewReply = "That's okay. Please continue."
      End Select
      Exit Do
    End If
    
  ' Check for "Dream" keyword
    If ((InStr(Question, "DREAM")) > 0) Then
      Choice = CInt(5 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "What does that dream suggest to you?"
        Case 2
          NewReply = "Do you dream often?"
        Case 3
          NewReply = "Are you disturbed by dreams?"
        Case 4
          NewReply = "Dreaming is natural."
        Case Else
          NewReply = "What do your dreams reveal about your thoughts?"
      End Select
      Exit Do
    End If
  
  ' Check for "Maybe" keyword
    If ((InStr(Question, "MAYBE")) > 0) Then
      Choice = CInt(6 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "You don't seem quite certain."
        Case 2
          NewReply = "Why the uncertain tone?"
        Case 3
          NewReply = "Can't you be more positive?"
        Case 4
          NewReply = "You aren't sure?"
        Case 5
          NewReply = "Don't you know?"
        Case Else
          NewReply = "Perhaps that might be true after all."
      End Select
      Exit Do
    End If
  
  ' Check for "Always" keyword
    If ((InStr(Question, "ALWAYS")) > 0) Then
      Choice = CInt(5 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Can you think of an example?"
        Case 2
          NewReply = "When?"
        Case 3
          NewReply = "What are you thinking about now?"
        Case 4
          NewReply = "Really...always?"
        Case Else
          NewReply = "Perhaps that might be true occasionally."
      End Select
      Exit Do
    End If
  
  ' Check for "Alike" keyword
    If ((InStr(Question, "ALIKE")) > 0) Then
      Choice = CInt(7 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "In what way?"
        Case 2
          NewReply = "What similarities do you see?"
        Case 3
          NewReply = "What does the similarity suggest to you?"
        Case 4
          NewReply = "Could there really be a connection?"
        Case 5
          NewReply = "How?"
        Case 6
          NewReply = "You seem quite positive."
        Case Else
          NewReply = "What other connections do you see?"
      End Select
      Exit Do
    End If
    
  ' Check for "Friend" keyword
    If ((InStr(Question, "FRIEND")) > 0) Then
      Choice = CInt(6 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Do you have many friends?"
        Case 2
          NewReply = "Do your friends worry you?"
        Case 3
          NewReply = "Do they pick on you?"
        Case 4
          NewReply = "Are your friends a source of anxiety?"
        Case 5
          NewReply = "Do you impose your problems on your friends?"
        Case Else
          NewReply = "Perhaps your dependency on your friends worries you."
      End Select
      Exit Do
    End If
    
  ' Check for "Computer" keyword
    If ((InStr(Question, "COMPUTER")) > 0) Or ((InStr(Question, "MACHINE")) > 0) Then
      Choice = CInt(6 * Rnd + 1)
      Select Case Choice
        Case 1
          NewReply = "Do computers worry you?"
        Case 2
          NewReply = "Are you talking about me in particular?"
        Case 3
          NewReply = "Do machines frighten you?"
        Case 4
          NewReply = "Why do you mention computers?"
        Case 5
          NewReply = "Don't you think computers can help you?"
        Case Else
          NewReply = "What is it about machines that worry you?"
      End Select
      Exit Do
    End If
  
    
  ' Check for "Can I" keyword
    If ((InStr(Question, "CAN I")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = ".") Or (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 5)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 5) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Do you want to be able to" & TempReply
        Case 2
          NewReply = "Perhaps you really don't want to" & TempReply
        Case 3
          NewReply = "What would it mean if you could" & TempReply
        Case Else
          NewReply = "Why do you ask if you can" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "Can you" keyword
    If ((InStr(Question, "CAN YOU")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = ".") Or (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 6)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 6) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Don't you believe that I can" & TempReply
        Case 2
          NewReply = "Perhaps you would like me to be able to" & TempReply
        Case 3
          NewReply = "You want me to be able to" & TempReply
        Case Else
          NewReply = "Why do you want me to" & TempReply
      End Select
      Exit Do
    End If
    
  ' Check for "Don't you" keyword
    If ((InStr(Question, "DON'T YOU")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = ".") Or (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 8)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 8) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Do you really believe that I don't" & TempReply
        Case 2
          NewReply = "Do you want me to" & TempReply
        Case 3
          NewReply = "Why do you care if I" & TempReply
        Case Else
          NewReply = "Is it important to you that I" & TempReply
      End Select
      Exit Do
    End If
    
  ' Check for "Are you" keyword
    If ((InStr(Question, "ARE YOU")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = ".") Or (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 6)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 6) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Would you prefer if I were not" & TempReply
        Case 2
          NewReply = "Perhaps you are imagining that I am" & TempReply
        Case 3
          NewReply = "Do you really care if I am" & TempReply
        Case Else
          NewReply = "Why should you care if I am" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "I don't" keyword
    If ((InStr(Question, "I DON'T")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 7)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 7) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Don't you really" & TempReply
        Case 2
          NewReply = "Why don't you" & TempReply
        Case 3
          NewReply = "Do you wish to be able to" & TempReply
        Case Else
          NewReply = "So you really don't" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "I feel" keyword
    If ((InStr(Question, "I FEEL")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 6)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 6) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Do you often feel" & TempReply
        Case 2
          NewReply = "Do you enjoy feeling" & TempReply
        Case 3
          NewReply = "How often do you feel" & TempReply
        Case Else
          NewReply = "What do you think about when you feel" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "I want" keyword
    If ((InStr(Question, "I WANT")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 6)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 6) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "What would it mean if you got" & TempReply
        Case 2
          NewReply = "Why do you want" & TempReply
        Case 3
          NewReply = "What would happen if you got" & TempReply
        Case Else
          NewReply = "What if you never got" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "I am" keyword
    If ((InStr(Question, "I AM")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 4)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 4) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Did you come to me because you are" & TempReply
        Case 2
          NewReply = "How long have you been" & TempReply
        Case 3
          NewReply = "Do you believe it is normal to be" & TempReply
        Case Else
          NewReply = "Why are you" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "I'm" keyword
    If ((InStr(Question, "I'M")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 3)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 3) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Did you come to me because you are" & TempReply
        Case 2
          NewReply = "How long have you been" & TempReply
        Case 3
          NewReply = "Do you believe it is normal to be" & TempReply
        Case Else
          NewReply = "Why are you" & TempReply
      End Select
      Exit Do
    End If
    
  ' Check for "I can't" keyword
    If ((InStr(Question, "I CAN'T")) = 1) Then
      Choice = CInt(4 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 7)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 7) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "How do you know you can't" & TempReply
        Case 2
          NewReply = "Perhaps now you can" & TempReply
        Case 3
          NewReply = "Why don't you try to" & TempReply
        Case Else
          NewReply = "How would you feel if you could" & TempReply
      End Select
      Exit Do
    End If
    
  ' Check for "You're" keyword
    If ((InStr(Question, "YOU'RE")) = 1) Then
      Choice = CInt(5 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 6)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 6) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Does it please you to believe that I am" & TempReply
        Case 2
          NewReply = "Perhaps you would like to be" & TempReply
        Case 3
          NewReply = "Do you sometimes wish you were" & TempReply
        Case 4
          NewReply = "What makes you think I am" & TempReply
        Case Else
          NewReply = "Why do you care that I might be" & TempReply
      End Select
      Exit Do
    End If
  
  ' Check for "You aren't" keyword
    If ((InStr(Question, "YOU AREN'T")) = 1) Then
      Choice = CInt(5 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 10)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 10) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Why do you think that I am not" & TempReply
        Case 2
          NewReply = "Is it important that you think I am not" & TempReply
        Case 3
          NewReply = "Why does it matter that I may not be" & TempReply
        Case 4
          NewReply = "Why do you think I am not" & TempReply
        Case Else
          NewReply = "Why do you care that I might not be" & TempReply
      End Select
      Exit Do
    End If
    
  ' Check for "You are" keyword
    If ((InStr(Question, "YOU ARE")) = 1) Then
      Choice = CInt(5 * Rnd + 1)
      ' Convert You to Me
      
      If (Right(Question, 1) = "?") Then
        TempReply = Right(LCase(Question), Len(Question) - 7)
      Else
        TempReply = Right(LCase(Question), Len(Question) - 7) & "?"
      End If
      Select Case Choice
        Case 1
          NewReply = "Does it please you to believe that I am" & TempReply
        Case 2
          NewReply = "Perhaps you would like to be" & TempReply
        Case 3
          NewReply = "Do you sometimes wish that you were" & TempReply
        Case 4
          NewReply = "What makes you think that I am" & TempReply
        Case Else
          NewReply = "Why do you care that I might be" & TempReply
      End Select
      Exit Do
    End If
    
  ' No keywords found
    Choice = CInt(8 * Rnd + 1)
    Select Case Choice
      Case 1
        NewReply = "Do you feel intense psychological stress right now?"
      Case 2
        NewReply = "What is your meaning of life?"
      Case 3
        NewReply = "You need alot of prayer!"
      Case 4
        NewReply = "Those who know won't tell but those who tell they really don't know anything."
      Case 5
        NewReply = "Have you ever been to a real psychiatrist?"
      Case 6
        NewReply = "That is quite interesting."
      Case 7
        NewReply = "It looks like I may need a Psychologist after this session!"
      Case Else
        NewReply = "I think you have homosexual tendencies!"
    End Select
    
    Loop Until NewReply <> ""
  Loop Until NewReply <> LastReply
End Function

