using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System;


string filePath = @"C:\source\TestFiles\MyPresentation.pptx";


using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
{

    if (presentationDocument.PresentationPart == null || presentationDocument.PresentationPart.Presentation.SlideIdList == null)
    {
        throw new NullReferenceException("Presenation Part is empty or there are no slides");
    }

    //Get presentation part
    PresentationPart presentationPart = presentationDocument.PresentationPart;

    //Get slides ids.
    OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;

    string startTransitionAfterMs = "6000", durationMs = "2000";
    //Set to false if you want to advance the slide on click
    bool advanceOnClick = true;


    foreach (SlideId slideId in slidesIds)
    {
        string? relId = slideId!.RelationshipId!.ToString();

        if (relId == null)
        {
            throw new NullReferenceException("RelationshipId not found");
        }

        SlidePart? slidePart = (SlidePart) presentationDocument.PresentationPart.GetPartById(relId);

        //Remove existing timing
        if (slidePart.Slide.Timing != null)
        {
            slidePart.Slide.Timing.Remove();
        }

        //Remove existing transitions
        if (slidePart.Slide.Transition != null)
        {
            slidePart.Slide.Transition.Remove();
        }

        if (slidePart!.Slide.Descendants<AlternateContent>().ToList().Count > 0)
        {
            List<AlternateContent> alternateContents = [.. slidePart.Slide.Descendants<AlternateContent>()];
            foreach (var alternateContent in alternateContents)
            {
                // remove transitions in AlternateContentChoice within AlternateContent
                var childElements = alternateContent.ChildElements.ToList();

                foreach (var element in childElements)
                {

                    var trans = element.Descendants<Transition>().ToList();
                    foreach (var transition in trans)
                    {
                        transition.Remove();
                    }
                }
            }

            foreach (var alternateContent in alternateContents)
            {
                alternateContent!.GetFirstChild<AlternateContentChoice>();
                alternateContent!.GetFirstChild<AlternateContentChoice>()!.Append(new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow });
                alternateContent!.GetFirstChild<AlternateContentFallback>()!.Append(new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow });
            }

        }
        //Add transition if there is none
        else
        {
            //Cheks if there is transition appended to the slide and set it to null
            if (slidePart.Slide.Transition != null)
            {
                slidePart.Slide.Transition = null;
            }

            AlternateContent alternateContent = new AlternateContent();
            alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "p14" };
            alternateContentChoice.Append(new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow });
            
            AlternateContentFallback alternateContentFallback = new AlternateContentFallback(new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = false, Speed = TransitionSpeedValues.Slow });
            alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main");
            alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative");
            alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main");

            alternateContent.Append(alternateContentChoice);
            alternateContent.Append(alternateContentFallback);
            slidePart.Slide.Append(alternateContent);
        }

    }

}
