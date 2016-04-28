/*
Tests comment stripping support in jet-tool.
Supported comments:
  /*  style possibly multiline comments
  --  style single line
Either is ignored inside the other.
*/

CREATE /* Comment */ TABLE /* Comment -- Comment */ [CommentsTest] /* Comment */ ( /* Comment */  -- Comment /*
[ID] COUNTER /* Comment */ NOT NULL, /* Comment */
[IntValue] INTEGER DEFAULT 0, /*

*/ [StringValue] /*

*/ TEXT(128)
) /* Comment */;
/* Comment */


/* Comment */