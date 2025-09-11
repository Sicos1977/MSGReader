using System;
using System.Collections.Generic;

namespace MsgReader.Mime.Traverse;

internal class TextVersionFinder : MultipleMessagePartFinder
{
    #region CaseLeaf
    protected override List<MessagePart> CaseLeaf(MessagePart messagePart)
    {
        if(messagePart == null)
            throw new ArgumentNullException(nameof(messagePart));

        var leafAnswer = new List<MessagePart>(1);

        if (messagePart.IsText)
            leafAnswer.Add(messagePart);

        return leafAnswer;
    }
    #endregion
}