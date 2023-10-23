
       // Declare a comments part.
            SlideCommentsPart commentsPart;

            // Verify that there is a comments part in the first slide part.
            if (slidePart1.GetPartsOfType<SlideCommentsPart>().Count() == 0)
            {
                // If not, add a new comments part.
                commentsPart = slidePart1.AddNewPart<SlideCommentsPart>();
            }
            else
            {
                // Else, use the first comments part in the slide part.
                commentsPart = slidePart1.GetPartsOfType<SlideCommentsPart>().First();
            }

            // If the comment list does not exist.
            if (commentsPart.CommentList == null)
            {
                // Add a new comments list.
                commentsPart.CommentList = new CommentList();
            }

            // Get the new comment ID.
            uint commentIdx = author.LastIndex == null ? 1 : author.LastIndex + 1;
            author.LastIndex = commentIdx;

            // Add a new comment.
            Comment comment = commentsPart.CommentList.AppendChild<Comment>(
                new Comment()
                {
                    AuthorId = authorId,
                    Index = commentIdx,
                    DateTime = DateTime.Now
                });

            // Add the position child node to the comment element.
            comment.Append(
                new Position() { X = 100, Y = 200 },
                new Text() { Text = text });

       // Declare a comments part.
            SlideCommentsPart commentsPart;

            // Verify that there is a comments part in the first slide part.
            if (slidePart1.GetPartsOfType<SlideCommentsPart>().Count() == 0)
            {
                // If not, add a new comments part.
                commentsPart = slidePart1.AddNewPart<SlideCommentsPart>();
            }
            else
            {
                // Else, use the first comments part in the slide part.
                commentsPart = slidePart1.GetPartsOfType<SlideCommentsPart>().First();
            }

            // If the comment list does not exist.
            if (commentsPart.CommentList == null)
            {
                // Add a new comments list.
                commentsPart.CommentList = new CommentList();
            }

            // Get the new comment ID.
            uint commentIdx = author.LastIndex == null ? 1 : author.LastIndex + 1;
            author.LastIndex = commentIdx;

            // Add a new comment.
            Comment comment = commentsPart.CommentList.AppendChild<Comment>(
                new Comment()
                {
                    AuthorId = authorId,
                    Index = commentIdx,
                    DateTime = DateTime.Now
                });

            // Add the position child node to the comment element.
            comment.Append(
                new Position() { X = 100, Y = 200 },
                new Text() { Text = text });