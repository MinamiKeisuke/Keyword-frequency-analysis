from flask import Flask, request, render_template, send_file
from janome.tokenizer import Tokenizer
import json
import os
from collections import Counter
import pandas as pd
from datetime import datetime

app = Flask(__name__)

USER_DEFINED_WORDS_FILE = 'user_defined_words.json'

def save_user_defined_words(new_words):
    if os.path.exists(USER_DEFINED_WORDS_FILE):
        try:
            with open(USER_DEFINED_WORDS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except json.JSONDecodeError:
            data = {"自立語": []}
    else:
        data = {"自立語": []}
    
    existing_words = set(data["自立語"])
    for word in new_words:
        existing_words.add(word.strip())
    
    data["自立語"] = list(existing_words)
    with open(USER_DEFINED_WORDS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def load_user_defined_words():
    if os.path.exists(USER_DEFINED_WORDS_FILE):
        try:
            with open(USER_DEFINED_WORDS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f).get("自立語", [])
        except json.JSONDecodeError:
            return []
    return []

def analyze_text(text, keywords, user_defined_words):
    tokenizer = Tokenizer()
    token_surfaces = []
    independent_words = []

    for word in user_defined_words:
        while word in text:
            independent_words.append(word)
            text = text.replace(word, "", 1)
    
    tokens = tokenizer.tokenize(text)
    total_independent_words = len(independent_words)
    
    independent_pos = ['名詞', '動詞', '形容詞', '形容動詞']
    
    for token in tokens:
        token_surfaces.append(token.surface)
        part_of_speech = token.part_of_speech.split(',')[0]
        if part_of_speech in independent_pos:
            total_independent_words += 1
            independent_words.append(token.surface)
    
    concatenated_text = ''.join(token_surfaces)

    keyword_counts = {}
    for keyword in keywords:
        keyword_counts[keyword] = concatenated_text.count(keyword) + independent_words.count(keyword)
    
    results = {}
    for keyword, count in keyword_counts.items():
        rate = round((count / total_independent_words) * 100, 2) if total_independent_words > 0 else 0
        remaining_to_3_percent = max(0, int((3 * total_independent_words / 100) - count))
        excess_over_8_percent = max(0, int(count - (8 * total_independent_words / 100)))
        results[keyword] = {
            'count': count,
            'rate': rate,
            'color': 'red' if rate < 3 or rate > 8 else 'black',
            'remaining_to_3_percent': remaining_to_3_percent,
            'excess_over_8_percent': excess_over_8_percent
        }

    top_words = Counter(independent_words).most_common(10)
    top_words_results = [
        {
            'word': word,
            'count': count,
            'rate': round((count / total_independent_words) * 100, 2)
        } for word, count in top_words
    ]
    
    return total_independent_words, independent_words, results, top_words_results

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        text = request.form['text']
        keywords = [keyword.strip() for keyword in request.form['keywords'].split(',')]
        new_user_defined_words = [word.strip() for word in request.form['new_user_defined_words'].split(',') if word.strip()]
        title = request.form['title']
        save_user_defined_words(new_user_defined_words)
        user_defined_words = load_user_defined_words()
        total_words, independent_words, analysis_results, top_words_results = analyze_text(text, keywords, user_defined_words)
        return render_template('index.html', total_words=total_words, independent_words=independent_words, results=analysis_results, top_words_results=top_words_results, text=text, keywords=keywords, user_defined_words=user_defined_words, title=title)
    else:
        user_defined_words = load_user_defined_words()
    return render_template('index.html', user_defined_words=user_defined_words)

@app.route('/download', methods=['POST'])
def download():
    text = request.form['text']
    keywords = [keyword.strip() for keyword in request.form['keywords'].split(',')]
    user_defined_words = load_user_defined_words()
    total_words, independent_words, analysis_results, top_words_results = analyze_text(text, keywords, user_defined_words)

    # DataFrames for Excel export
    df_keywords = pd.DataFrame([
        {
            'キーワード': keyword,
            '回数': data['count'],
            '出現率 (%)': data['rate'],
            '追加': data['remaining_to_3_percent'] if data['rate'] < 3 else '',
            '削除': data['excess_over_8_percent'] if data['rate'] > 8 else ''
        } for keyword, data in analysis_results.items()
    ])
    
    df_top_words = pd.DataFrame([
        {
            '自立語': word['word'],
            '回数': word['count'],
            '出現率 (%)': word['rate']
        } for word in top_words_results
    ])

    # Create a Pandas Excel writer using openpyxl as the engine.
    with pd.ExcelWriter('/tmp/analysis_results.xlsx', engine='openpyxl') as writer:
        df_keywords.to_excel(writer, sheet_name='Keywords Analysis', index=False)
        df_top_words.to_excel(writer, sheet_name='Top 10 Words', index=False)

    # Get the title or keywords for the filename
    title = request.form['title']
    date_str = datetime.now().strftime('%Y-%m-%d')
    if title:
        filename = f"{title}-{date_str}.xlsx"
    else:
        filename = f"{'-'.join(keywords)}-{date_str}.xlsx"

    return send_file('/tmp/analysis_results.xlsx', as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(debug=True)
