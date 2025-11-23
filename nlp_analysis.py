from bertopic import BERTopic
from umap.umap_ import UMAP

def perform_nlp_analysis(preprocessed_texts, n_topics=50, min_topic_size=20):
    umap_model = UMAP(n_neighbors=15, min_dist=0.1, metric='cosine')
    topic_model = BERTopic(umap_model=umap_model, nr_topics=n_topics, min_topic_size=min_topic_size)
    topics, _ = topic_model.fit_transform(preprocessed_texts)
    return topic_model, topics

def create_topics(preprocessed_texts, nummer):
    topic_model, topics = perform_nlp_analysis(preprocessed_texts)
    topic_to_inc = {}
    for topic, num in zip(topics, nummer):
        if topic not in topic_to_inc:
            topic_to_inc[topic] = []
        topic_to_inc[topic].append(num)
    return topic_model, topic_to_inc