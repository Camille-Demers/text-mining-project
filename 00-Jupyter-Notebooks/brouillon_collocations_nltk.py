# https://www.nltk.org/howto/collocations.html 

from nltk.collocations import *
bigram_measures = nltk.collocations.BigramAssocMeasures()
trigram_measures = nltk.collocations.TrigramAssocMeasures()
fourgram_measures = nltk.collocations.QuadgramAssocMeasures()

# Bigrammes
finder_b = BigramCollocationFinder.from_words(tokens)
finder_b.apply_freq_filter(10)
finder_b.apply_word_filter(lambda w : w in stopwords or len(w) < 3)

# Trigrammes
finder_t = TrigramCollocationFinder.from_words(tokens)
finder_t.apply_freq_filter(10)
finder_t.apply_ngram_filter(lambda w1, w2, w3 : w1 in stopwords or len(w1) < 3 or w3 in stopwords or len(w3) < 3)

# Quadgrammes
finder_q = QuadgramCollocationFinder.from_words(tokens)
finder_q.apply_freq_filter(10)
finder_q.apply_ngram_filter(lambda w1, w2, w3, w4 : w1 in stopwords or len(w1) < 3 or w4 in stopwords or len(w4) < 3)

lr_b = finder_b.score_ngrams(bigram_measures.likelihood_ratio)
lr_t = finder_t.score_ngrams(trigram_measures.likelihood_ratio)
lr_q = finder_q.score_ngrams(fourgram_measures.likelihood_ratio)

output_path = '../04-filtrage/output/' 
output_path = path.join(output_path, acteur, acteur + '-LLR_bigrams.txt')
print(output_path)
with open(output_path, 'w', encoding='utf-8') as f: 
    f.write("Collocation\tScore ({})".format(bigram_measures.likelihood_ratio.__name__) + '\n ')
    for (ngram), score in lr_b: f.write("{}\t{}\n".format(repr(ngram), score))

output_path = '../04-filtrage/output/' 
output_path = path.join(output_path, acteur, acteur + '-LLR_trigrams.txt')
print(output_path)
with open(output_path, 'w', encoding='utf-8') as f: 
    f.write("Collocation\tScore ({})".format(trigram_measures.likelihood_ratio.__name__) + '\n ')
    for (ngram), score in lr_t: f.write("{}\t{}\n".format(repr(ngram), score))

output_path = '../04-filtrage/output/' 
output_path = path.join(output_path, acteur, acteur + '-LLR_quadgrams.txt')
print(output_path)
with open(output_path, 'w', encoding='utf-8') as f: 
    f.write("Collocation\tScore ({})".format(fourgram_measures.likelihood_ratio.__name__) + '\n ')
    for (ngram), score in lr_q: f.write("{}\t{}\n".format(repr(ngram), score))

