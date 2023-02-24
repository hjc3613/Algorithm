import spacy
from spacy.matcher import Matcher
from spacy.util import filter_spans

# nlp = spacy.load('zh_core_web_sm')
nlp = spacy.blank('zh')

patterns = [
    [{'TEXT':'北'}, {'IS_SPACE':True}, {'TEXT':'京'}],
    [{'TEXT':'北'}, {'TEXT':'京'}]
]
matcher = Matcher(nlp.vocab)
matcher.add('北京', patterns)
doc = nlp('我北京爱你, 北京！')
matches = matcher(doc)
for match_id, start, end in matches:
    string_id = nlp.vocab.strings[match_id]  # Get string representation
    span = doc[start:end]  # The matched span
    print(match_id, string_id, start, end, span.text)