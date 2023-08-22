from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    sections = {
        "Introduction": "1-23",
        "WPD": "24-34",
        "AMD": "35-41",
        "Compliance": "42-44",
        "Marketing": "45-49",
        "Opportunities": "50-53",
        "Important Facts": "54-61",
        "Our Capital Partners": "62-65",
        "Vision 2030": "66-70",
        "ALL": "1-70"
    }

    if request.method == 'POST':
        user_name = request.form['user_name']
        selected_sections = request.form.getlist('sections')

        # Now, call your convert_ppt_to_pdf function here with these details
        # convert_ppt_to_pdf(ppt_file, output_folder, selected_sections, user_name)

        return redirect(url_for('success'))

    return render_template('index.html', sections=sections)

@app.route('/success')
def success():
    return "Information collected successfully!"

if __name__ == '__main__':
    app.run(debug=True)
