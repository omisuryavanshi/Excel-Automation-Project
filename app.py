
import os
import uuid
import pandas as pd

from flask import (
    Flask,
    render_template,
    request,
    send_from_directory,
    url_for
)

from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/process", methods=["POST"])
def process_excel():

    file = request.files.get("excel_file")

    if not file or file.filename == "":
        return render_template(
            "index.html",
            message="Please select an Excel file."
        )

    if not file.filename.lower().endswith(".xlsx"):
        return render_template(
            "index.html",
            message="Please upload an .xlsx file."
        )

    filename = secure_filename(file.filename)

    unique_name = uuid.uuid4().hex
    input_path = os.path.join(
        UPLOAD_FOLDER,
        unique_name + "_" + filename
    )

    file.save(input_path)

    try:
        df = pd.read_excel(input_path)

        required_columns = [
            "Customer", "Product", "Quantity", "Price"
        ]

        if not all(col in df.columns for col in required_columns):
            raise ValueError("Excel file must contain Customer, Product, Quantity and Price columns.")

        df = df.drop_duplicates()
        df = df.dropna(subset=required_columns)

        df["Total Sales"] = df["Quantity"] * df["Price"]

        report = df.groupby("Product")["Total Sales"].sum()

        total_revenue = df["Total Sales"].sum()
        total_quantity = df["Quantity"].sum()
        total_customers = df["Customer"].nunique()

        best_product = (
            df.groupby("Product")["Quantity"]
            .sum()
            .idxmax()
        )

        average_sales = df["Total Sales"].mean()

        summary = pd.DataFrame({
            "Metric": [
                "Total Revenue",
                "Total Products Sold",
                "Total Unique Customers",
                "Best Selling Product",
                "Average Sales per Row"
            ],
            "Value": [
                total_revenue,
                total_quantity,
                total_customers,
                best_product,
                round(average_sales, 2)
            ]
        })

        report_name = unique_name + "_product_report.xlsx"
        summary_name = unique_name + "_sales_summary.xlsx"
        cleaned_name = unique_name + "_cleaned_sales.xlsx"

        df.to_excel(
            os.path.join(OUTPUT_FOLDER, cleaned_name),
            index=False
        )

        report.to_excel(
            os.path.join(OUTPUT_FOLDER, report_name)
        )

        summary.to_excel(
            os.path.join(OUTPUT_FOLDER, summary_name),
            index=False
        )

        return render_template(
            "index.html",
            message="Your Excel file was processed successfully!",
            files=[
                {"name": cleaned_name},
                {"name": report_name},
                {"name": summary_name}
            ],
            total_revenue=total_revenue,
            total_quantity=total_quantity,
            total_customers=total_customers,
            best_product=best_product
        )

    except Exception as error:
        return render_template(
            "index.html",
            message=f"Error processing file: {error}"
        )

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)


@app.route("/download/<filename>")
def download_file(filename):

    safe_name = secure_filename(filename)

    if safe_name != filename:
        return "Invalid filename", 400

    return send_from_directory(
        OUTPUT_FOLDER,
        safe_name,
        as_attachment=True
    )

# History Page
@app.route("/history")
def history_page():
    return render_template("history.html", history=[])


# Settings Page
@app.route("/settings")
def settings_page():
    return render_template("settings.html")


if __name__ == "__main__":
    app.run(debug=True)


