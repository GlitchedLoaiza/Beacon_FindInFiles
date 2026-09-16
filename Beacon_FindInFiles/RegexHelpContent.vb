Namespace Beacon
    Public NotInheritable Class RegexHelpExample
        Public ReadOnly Property Title As String
        Public ReadOnly Property Pattern As String
        Public ReadOnly Property Input As String
        Public ReadOnly Property Expected As String
        Public ReadOnly Property NonMatchingInput As String
        Public ReadOnly Property Explanation As String

        Public Sub New(title As String, pattern As String, input As String, expected As String, nonMatchingInput As String, explanation As String)
            Me.Title = title
            Me.Pattern = pattern
            Me.Input = input
            Me.Expected = expected
            Me.NonMatchingInput = nonMatchingInput
            Me.Explanation = explanation
        End Sub

        Public Function Description() As String
            Return Title & vbCrLf & "Enter: " & Quote(Pattern) & vbCrLf & "In: " & Quote(Input) & vbCrLf &
                "Finds/highlights: " & Quote(Expected) & vbCrLf & "Does not match: " & Quote(NonMatchingInput) & vbCrLf & Explanation
        End Function

        Private Shared Function Quote(value As String) As String
            Return """" & value & """"
        End Function
    End Class

    Public NotInheritable Class RegexHelpContent
        Public Const DocumentationUrl As String = "https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference"
        Private Sub New()
        End Sub

        Public Shared Function Examples() As List(Of RegexHelpExample)
            Return New List(Of RegexHelpExample) From {
                New RegexHelpExample("Find changing numeric codes", "code=\d+", "Request failed: code=503", "code=503", "code=unknown", "The letters and equals sign are literal. \d means a decimal digit and + means one or more. This finds the code; it does not limit its length."),
                New RegexHelpExample("Find either of two words", "\b(error|failed)\b", "The operation FAILED today", "FAILED", "The operation failedness", "Parentheses group alternatives; | means OR. \b keeps each alternative at a word boundary. With Case sensitive unchecked, FAILED also matches failed."),
                New RegexHelpExample("Find one word, not a longer word", "\btimeout\b", "A timeout occurred", "timeout", "timeouts occurred", "\b marks a word boundary without including spaces in the highlight. For a single ordinary word you can use Whole word mode instead."),
                New RegexHelpExample("Find a literal dotted address", "192\.168\.1\.10\b", "Server 192.168.1.10 is unreachable", "192.168.1.10", "Server 192x168x1x10 is unreachable", "A plain dot means any character. Backslash-dot means an actual dot. This locates that address text; it is not a general IP-address validator."),
                New RegexHelpExample("Find flexible spacing", "status\s*[:=]\s*5[0-9]{2}\b", "status : 503", "status : 503", "status=404", "\s* allows zero or more whitespace characters, [: =] without the space is written [:=] to allow colon OR equals, and 5[0-9]{2} finds 500–599. Whitespace can include newlines inside an event/request; use [ \t]* if only spaces and tabs are wanted."),
                New RegexHelpExample("Find a line starting with ERROR", "^ERROR\b.*", "ERROR disk full", "ERROR disk full", "INFO ERROR disk full", "^ means the start of a line in Beacon. .* reads the rest of the line (except the newline). Lines with a leading timestamp or space will not match this pattern."),
                New RegexHelpExample("Find a complete date-shaped line", "^[0-9]{4}-[0-9]{2}-[0-9]{2}\r?$", "2026-01-02", "2026-01-02", "Date: 2026-01-02", "^ and $ anchor the pattern to the line. \r? allows the optional carriage return in Windows line endings. This checks shape only: an impossible date such as 2026-99-99 also has this shape. Use date filters when filtering event/request timestamps."),
                New RegexHelpExample("Find a short request identifier", "\brequestId=[A-Za-z0-9_-]+", "requestId=AB_42-x completed", "requestId=AB_42-x", "requestId= completed", "A character class lists allowed characters. Here letters, ASCII digits, underscore and hyphen are allowed, with at least one required. The hyphen is literal because it is last in the class."),
                New RegexHelpExample("Find two words in a particular order", "error[^\r\n]*timeout", "error while waiting: timeout", "error while waiting: timeout", "timeout happened before error", "[^\r\n]* reads characters other than line endings. It requires error before timeout on the same line. Use All terms instead when order should not matter."),
                New RegexHelpExample("Allow an optional letter", "\bcolou?r\b", "The colour changed", "colour", "The colourful panel", "? makes only the preceding item optional: this matches color or colour, but not colourful."),
                New RegexHelpExample("Find text inside the first pair of brackets", "\[[^\]\r\n]+\]", "[ERROR] request failed", "[ERROR]", "ERROR request failed", "\[ and \] are literal brackets. The negated class between them accepts one or more characters other than a closing bracket or a line ending."),
                New RegexHelpExample("Highlight only digits after a label (advanced)", "(?<=code=)[0-9]+", "code=503", "503", "status=503", "This positive lookbehind requires code= immediately before the digits, but highlights only the digits. Lookbehind is optional advanced syntax, not needed for your first searches.")
            }
        End Function

        Public Shared Function Topics() As List(Of HelpTopic)
            Dim lessons As New List(Of HelpTopic) From {
                New HelpTopic("Regex 1 — your first pattern",
                    "WHAT IS REGEX?" & vbCrLf & "A regular expression (regex) is a description of text you want to find. Ordinary search asks for one exact value. Regex can find many values with the same shape. For example, you may want every numeric error code rather than search for each number separately. Beacon uses Microsoft's .NET regex syntax. Regex only searches here; it does not replace text or edit your files.",
                    "TRY THIS FIRST" & vbCrLf & "1. Choose a folder or archive that contains logs." & vbCrLf & "2. Select Regular expression in Search mode." & vbCrLf & "3. Enter ""code=\d+"" in Search for. Type one backslash. Do NOT include the example's surrounding double quotes, slash delimiters, a language prefix or flags." & vbCrLf & "4. Leave Case sensitive unchecked and click Scan." & vbCrLf & "5. Select a result and inspect Match details/the preview.",
                    "WHAT TO EXPECT" & vbCrLf & "If a searched record contains ""code=503"", the pattern matches and that text can be highlighted. It also finds ""code=7"" and ""CODE=42"" when case sensitivity is off. It does not find ""code=unknown"" or ""code = 503"" because the first pattern expects the equals sign without spaces. No matching records means no results for that pattern; a pattern does not create sample data.",
                    "READ IT PIECE BY PIECE" & vbCrLf & """code="" means those exact characters. ""\d"" means one decimal digit. ""+"" means repeat the preceding item one or more times. Put together, it means: find code= followed by at least one digit.",
                    "QUOTES AND BACKSLASHES" & vbCrLf & "Surrounding double quotes in this guide label examples: leave them out. In Regex mode, any quote you type is part of the pattern and normally searches for a quote character. Unlike Any term/All terms, quotes do not group phrases. Backslashes ARE part of regex syntax: type ""\d"", not ""\\d"". The latter searches for a literal backslash followed by d. Do not wrap a pattern as /pattern/ or add /i or /g; use Beacon's Case sensitive checkbox.",
                    "WHEN NOT TO USE REGEX" & vbCrLf & "For a fixed word, phrase, path or punctuation use Literal text. Use Whole word to avoid matching parts of words. Use Any term or All terms for simple word combinations. Regex is useful when those choices cannot express what you need."),
                New HelpTopic("Regex 2 — symbols explained",
                    "HOW TO READ THIS LESSON" & vbCrLf & "Each symbol is shown inside double quotes for readability; do not type the surrounding quotes. A pattern is assembled left to right. Most letters and numbers match themselves. Spaces you type normally matter.",
                    "CHARACTERS" & vbCrLf & """."" = any one character except a newline. ""\."" = an actual dot. ""\d"" = a Unicode decimal digit; ""[0-9]"" restricts it to ASCII digits. ""\w"" includes letters, digits, underscore and other Unicode word characters—not just English letters. ""\s"" = whitespace, including spaces, tabs and newlines. ""\t"" = tab; ""\r"" = carriage return; ""\n"" = newline. ""[ \t]"" accepts a space or tab without crossing a line.",
                    "CHARACTER CLASSES" & vbCrLf & """[ABC]"" means one A, B or C. ""[0-9]"" means one digit from 0 to 9. ""[A-Za-z]"" means one ASCII letter. ""[^,]"" means one character that is NOT a comma; it can include a newline. The ^ negates only when it appears at the beginning inside brackets. Put a literal hyphen last or escape it so it is not mistaken for a range.",
                    "HOW MANY?" & vbCrLf & """+"" = one or more. ""*"" = zero or more. ""?"" = zero or one (optional). ""{3}"" = exactly three. ""{2,4}"" = two through four. These affect the preceding character/class/group, not the entire pattern. Example: ""[0-9]{3}"" finds three digits, possibly inside a longer number; use boundaries when that matters.",
                    "GROUPS AND CHOICES" & vbCrLf & "Parentheses group a part of the pattern. ""(error|warning)"" means error OR warning. ""(?:error|warning)"" makes the same choice without saving a separate capture. Captures are useful in programming, but Beacon highlights the whole match, not just a capture group. Use parentheses to keep choices together: ""status=(400|500)"".",
                    "POSITIONS, NOT LETTERS" & vbCrLf & """^"" = beginning of a line. ""$"" = end of a line in Beacon. ""\b"" = word boundary, for example ""\bcat\b"". An underscore is a word character, so this does not match cat inside ""cat_name"". ""\A"" and ""\z"" mark the start/end of the entire record being searched, not necessarily the entire file.",
                    "LITERAL PUNCTUATION" & vbCrLf & "Outside character classes, characters such as . + * ? ( ) [ ] { } ^ $ | and backslash have special roles. Put a backslash before a special character when you want its literal meaning: ""\[ERROR\]"" finds ""[ERROR]"". To match one literal backslash, the pattern needs two: ""\\"". Slash / is not a .NET regex delimiter in Beacon and normally needs no escaping."),
                New HelpTopic("Regex 3 — build a useful log search",
                    "GOAL: FIND SERVER ERROR STATUS CODES" & vbCrLf & "Suppose logs contain ""status=503"", ""status : 500"" and ""status=200"". We want the first two but not the successful 200 response. Choose Regular expression and build the pattern in small steps.",
                    "1. START WITH A KNOWN LABEL" & vbCrLf & """status"" finds that word anywhere. ""status="" requires an equals sign immediately after it. If the logs use different spacing/punctuation, that exact version will miss them.",
                    "2. ALLOW SPACING AND EITHER SEPARATOR" & vbCrLf & """status[ \t]*[:=][ \t]*"" means status, optional spaces/tabs, colon OR equals, then optional spaces/tabs. Each * allows zero characters, so no spaces is also valid. The [] around := means a choice of one character, not both together.",
                    "3. ADD THE NUMBER SHAPE" & vbCrLf & """5[0-9]{2}"" means 5 followed by exactly two digits. Add ""\b"" after it to avoid accepting the beginning of a longer word/number. The finished input is ""status[ \t]*[:=][ \t]*5[0-9]{2}\b"". It matches ""status : 503"" and ""status=500"", not ""status=200"" or ""status=5030"".",
                    "4. RUN AND CHECK" & vbCrLf & "Click Scan and inspect both matches and Diagnostics. If nothing matches, first try Literal text with ""status"" to confirm the selected files contain the label. Then check the actual spelling and punctuation before making the pattern more complicated. For HAR status fields, the dedicated HAR status filter can be simpler and more precise than a text pattern.",
                    "GREEDY VERSUS LAZY" & vbCrLf & "A repetition such as * or + normally takes as much as it can while allowing the pattern to succeed. ""\[.*\]"" can match all of ""[first] [second]"". Adding ? makes it lazy: ""\[.*?\]"" finds ""[first]"" and then ""[second]"" as separate matches. Prefer a class excluding the delimiter, such as ""\[[^\]\r\n]+\]"", when it describes the data more clearly.",
                    "REQUIRE CONTEXT WITHOUT HIGHLIGHTING IT (ADVANCED)" & vbCrLf & """(?<=code=)[0-9]+"" finds just ""503"" in ""code=503"". The lookbehind checks what precedes the number but does not include it in the highlight. ""error(?=:)"" matches error only if a colon follows. These are optional tools; a normal pattern that includes the label is usually easier to understand."),
                New HelpTopic("Regex 4 — common recipes",
                    "HOW TO USE THESE" & vbCrLf & "Choose Regular expression. Copy only the characters between the example-label double quotes after Enter; do not include the outside quotes. The examples assume Case sensitive is unchecked. The displayed source text is illustrative, not data Beacon inserts into your files. Each recipe states exactly what it finds; modify cautiously and test on known data.",
                    String.Join(vbCrLf & vbCrLf, Examples().Select(Function(example) example.Description())),
                    "IMPORTANT" & vbCrLf & "A date-shaped pattern is not a calendar validator, an address-text pattern is not a complete network-address validator, and a text pattern cannot replace structured EVTX/HAR filters. Prefer the built-in ID, status and UTC filters when those are the fields you need."),
                New HelpTopic("Regex 5 — troubleshooting and boundaries",
                    "NO RESULTS?" & vbCrLf & "Check Regular expression is selected, the source contains your sample, and Case sensitive is correct. Leave out the guide's outer quotes and any /pattern/ delimiters. Check literal punctuation and spaces. Try a smaller part of the pattern first; then add one piece at a time.",
                    "INVALID PATTERN?" & vbCrLf & "Unclosed parentheses or brackets and a dangling backslash are common errors. ""["" is incomplete; ""\["" searches for a literal opening bracket. To find ""C:\Logs"" with regex, enter ""C:\\Logs"". Literal text is easier if you simply want that path. Beacon reports invalid patterns before scanning.",
                    "WHAT DOES BEACON SEARCH AT ONCE?" & vbCrLf & "Plain text is searched one line at a time. EVTX searches the combined event fields/message, with XML as a fallback; HAR searches combined request/response fields for each request. HTML/XML/JSON file search uses extracted document text. A regex cannot join two separate text records, events, requests or files. A preview may have different formatting, and privacy settings may hide the matching HAR value.",
                    "MULTILINE DOES NOT MEAN DOT MATCHES EVERYTHING" & vbCrLf & "Beacon enables .NET Multiline: ^ and $ refer to line starts/ends inside the searched record. Dot does not match newline by default. An advanced pattern can use ""(?s)"" to allow dot across newlines inside ONE record, for example ""(?s)error.*timeout"". It still cannot cross from one event/request or plain-text line into the next. For a Windows CRLF ending, ""\r?$"" allows the optional carriage return before the newline.",
                    "A MATCH WITHOUT A HIGHLIGHT?" & vbCrLf & "A pattern such as ""(?=error)"" matches a position before error, not the letters. Beacon can include the record but has no characters to highlight. Use ""error"" if you want the word highlighted. Parentheses alone do not limit highlighting to a capture group; the whole regex match is highlighted.",
                    "TOO MUCH MATCHED?" & vbCrLf & "Check whether a dot should be escaped, whether * should be +, and whether repetition should stop at a delimiter or line ending. Avoid assuming that ""[0-9]{3}"" excludes longer numbers unless you also constrain the boundaries. Case sensitive also applies to regex unless you deliberately override it with .NET inline options such as ""(?i)"".",
                    "TIMEOUT OR PARTIAL RESULTS?" & vbCrLf & "Beacon limits each regex operation to 100 milliseconds. Nested repetitions and many overlapping alternatives can cause excessive backtracking; avoid patterns such as ""(a+)+$"" on large input. Prefer clear labels, bounded character classes and modest repetition. Simplify the expression rather than trying to disable safeguards. Queries are limited to 4,096 characters and visible highlights to 10,000 examined matches. Scan/preview limits and first-match settings still apply.",
                    "OFFICIAL REFERENCE" & vbCrLf & "Use the Microsoft .NET regex documentation link below for the authoritative syntax reference and links to detailed explanations of character classes, anchors, groups and quantifiers. The guide here works offline; the external Microsoft page requires internet access.")
            }
            For Each topic In lessons
                topic.DocumentationUrl = DocumentationUrl
            Next
            Return lessons
        End Function
    End Class
End Namespace
