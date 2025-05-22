using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System;

// <Snippet0>
AddTransmitionToSlides(args[0]);
static void AddTransmitionToSlides(string filePath)
{
    // <Snippet1>
    using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
    // </Snippet1>
    {
    
     // Check if the presentation part and slide list are available
        if (presentationDocument.PresentationPart == null || presentationDocument.PresentationPart.Presentation.SlideIdList == null)
        {
            throw new NullReferenceException("Presentation part is empty or there are no slides");
        }

        // Get the presentation part
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the list of slide IDs
        OpenXmlElementList slidesIds = presentationPart.Presentation.SlideIdList.ChildElements;

        // <Snippet2>
        // Define the transition start time and duration in milliseconds
        string startTransitionAfterMs = "3000", durationMs = "2000";

        // Set to true if you want to advance to the next slide on mouse click
        bool advanceOnClick = true;
     
        // Iterate through each slide ID to get slides parts
        foreach (SlideId slideId in slidesIds)
        {
            // Get the relationship ID of the slide
            string? relId = slideId!.RelationshipId!.ToString();

            if (relId == null)
            {
                throw new NullReferenceException("RelationshipId not found");
            }

            // Get the slide part using the relationship ID
            SlidePart? slidePart = (SlidePart)presentationDocument.PresentationPart.GetPartById(relId);

            // Remove existing transitions if any
            if (slidePart.Slide.Transition != null)
            {
                slidePart.Slide.Transition.Remove();
            }

            // Check if there are any AlternateContent elements
            if (slidePart!.Slide.Descendants<AlternateContent>().ToList().Count > 0)
            {
                // Get all AlternateContent elements
                List<AlternateContent> alternateContents = [.. slidePart.Slide.Descendants<AlternateContent>()];
                foreach (AlternateContent alternateContent in alternateContents)
                {
                    // Remove transitions in AlternateContentChoice within AlternateContent
                    List<OpenXmlElement> childElements = alternateContent.ChildElements.ToList();

                    foreach (OpenXmlElement element in childElements)
                    {
                        List<Transition> transitions = element.Descendants<Transition>().ToList();
                        foreach (Transition transition in transitions)
                        {
                            transition.Remove();
                        }
                    }
                    // Add new transitions to AlternateContentChoice and AlternateContentFallback
                    alternateContent!.GetFirstChild<AlternateContentChoice>();
                    Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    Transition fallbackTransition = new Transition(new RandomBarTransition()) {AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                    alternateContent!.GetFirstChild<AlternateContentChoice>()!.Append(choiceTransition);
                    alternateContent!.GetFirstChild<AlternateContentFallback>()!.Append(fallbackTransition);
                }
            }
            // </Snippet2>
            
            // <Snippet3>
            // Add transition if there is none
            else
            {
                // Check if there is a transition appended to the slide and set it to null
                if (slidePart.Slide.Transition != null)
                {
                    slidePart.Slide.Transition = null;
                }
                // Create a new AlternateContent element
                AlternateContent alternateContent = new AlternateContent();
                alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                // Create a new AlternateContentChoice element and add the transition
                AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "p14" };
                Transition choiceTransition = new Transition(new RandomBarTransition()) { Duration = durationMs, AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                Transition fallbackTransition = new Transition(new RandomBarTransition()) { AdvanceAfterTime = startTransitionAfterMs, AdvanceOnClick = advanceOnClick, Speed = TransitionSpeedValues.Slow };
                alternateContentChoice.Append(choiceTransition);

                // Create a new AlternateContentFallback element and add the transition
                AlternateContentFallback alternateContentFallback = new AlternateContentFallback(fallbackTransition);
                alternateContentFallback.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                alternateContentFallback.AddNamespaceDeclaration("p16", "http://schemas.microsoft.com/office/powerpoint/2015/main");
                alternateContentFallback.AddNamespaceDeclaration("adec", "http://schemas.microsoft.com/office/drawing/2017/decorative");
                alternateContentFallback.AddNamespaceDeclaration("a16", "http://schemas.microsoft.com/office/drawing/2014/main");

                // Append the AlternateContentChoice and AlternateContentFallback to the AlternateContent
                alternateContent.Append(alternateContentChoice);
                alternateContent.Append(alternateContentFallback);
                slidePart.Slide.Append(alternateContent);
            } 
         // </Snippet3>
        }
    }
}
// </Snippet0>
