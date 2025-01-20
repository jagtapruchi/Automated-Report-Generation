import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
import matplotlib.pyplot as plt
import csv

#Function to create clean tables
def create_clean_table(data, col_widths=None):
    table = Table(data, colWidths=col_widths)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ])
    table.setStyle(style)
    return table

#Read sales data from CSV
sales_data = []
file_path = "C:/Users/RUCHI/Codetech IT/salesreport.csv"
with open(file_path, "r") as file:
    reader = csv.reader(file, delimiter="\t")  # Use tab delimiter for proper parsing
    for row in reader:
        sales_data.append(row)

#Create Discount vs Total Sales Bar Graph
def create_discount_vs_sales_chart(sales_data, file_name):
    discounts = ['0%', '5%', '10%', '15%']
    sales_by_discount = [0, 0, 0, 0]

    for i in range(1, len(sales_data)):
       try:
            row = sales_data[i]
            if len(row) < 8:  # Ensure the row has all columns
                print(f"Skipping malformed row: {row}")
                continue
            discount = row[7]
            sales = float(row[4].replace(",", ""))
            if discount == '0%':
                sales_by_discount[0] += sales
            elif discount == '5%':
                sales_by_discount[1] += sales
            elif discount == '10%':
                sales_by_discount[2] += sales
            elif discount == '15%':
                sales_by_discount[3] += sales
       except ValueError as e:
           print(f"Error processing row {row}: {e}")
       except IndexError as e:
           print(f"Row index error: {row} - {e}")

    plt.figure(figsize=(8, 6))
    plt.bar(discounts, sales_by_discount, color='salmon')
    plt.title('Discount Applied vs Total Sales')
    plt.xlabel('Discount')
    plt.ylabel('Total Sales')
    plt.tight_layout()
    plt.savefig('discount_sales_bargraph.png')
    plt.close()

create_discount_vs_sales_chart(sales_data, "C:/Users/RUCHI/Codetech IT/discount_sales_bargraph.png")

# Sales Trends Pie Chart
de = pd.read_csv("C:/Users/RUCHI/Codetech IT/salesreport.csv", delimiter="\t")

if de['Total Sales'].dtype != 'object':  
    de['Total Sales'] = de['Total Sales'].astype(str)

de['Total Sales'] = de['Total Sales'].str.replace(",", "").astype(float)  

sales_by_product = de.groupby('Product')['Total Sales'].sum()

plt.figure(figsize=(8, 8))
plt.pie(sales_by_product, labels=sales_by_product.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
plt.title('Sales Trends by Product')
plt.axis('equal')  
plt.tight_layout()

chart_path = "C:/Users/RUCHI/Codetech IT/sales_trends.png"
plt.savefig(chart_path)
plt.close()

#PDF Creation
pdf_path = "C:/Users/RUCHI/Codetech IT/clean_sales_report.pdf"
doc = SimpleDocTemplate(pdf_path, pagesize=letter,rightMargin=50, leftMargin=50,topMargin=50,bottomMargin=50)

#Title,Body and Bullet points Styling
styles = getSampleStyleSheet()
title_style = ParagraphStyle(name="Title", fontSize=19, leading=22, alignment=TA_CENTER, spaceAfter=30, fontName="Helvetica-Bold")
title = Paragraph("January Sales Report of EverPeak Co.Ltd.", title_style)
elements = [title]

body_style = ParagraphStyle(name="Body",fontSize=12,leading=16,alignment=TA_JUSTIFY,spaceAfter=14,fontName="Helvetica")

bullet_style = ParagraphStyle(name="Bullet",fontSize=12,leading=16,alignment=TA_JUSTIFY,bulletIndent=15,spaceAfter=8,fontName="Helvetica")

#Introductory section:
introduction = """
The January 2025 Sales Report offers a comprehensive analysis of sales performance over the first 12 days of the year, providing critical insights into product performance, regional sales distribution, and the effectiveness of sales strategies. This report consolidates transactional data across Maharashtra, detailing metrics such as units sold, revenue generated, discounts applied, and payment methods utilized.

The report covers five key products—Product A, B, C, D, and E—and highlights their contributions to overall sales. It also evaluates the performance of individual sales representatives and their impact on sales across regions, offering a detailed breakdown of market preferences and customer buying behaviors.
"""
elements.append(Paragraph(introduction,body_style))

elements.append(Paragraph("Key objectives of this report include:",body_style))
objectives = [
    "Identifying top-performing products, regions, and sales personnel.","Analyzing revenue trends and discount strategies to evaluate their influence on customer purchase decisions.",
    "Examining payment preferences to refine future payment offerings and enhance customer satisfaction.",
    "Providing actionable insights to drive strategic planning, optimize sales efforts, and improve revenue growth."]
for obj in objectives:
    elements.append(Paragraph(f"• {obj}",bullet_style))

closing = """
    By leveraging this report, stakeholders will gain valuable data-driven insights to assess current sales strategies, refine operational efficiencies, and establish a roadmap for sustained growth in the months ahead. This document serves as an essential tool for informed decision-making and performance benchmarking.
    """
elements.append(Paragraph(closing,body_style))
elements.append(Spacer(1,10))

#Add tables
productwise_data = [
    ["Product", "Units Sold", "Total Sales ", "Average Unit Price "],
    ["Product A", 300, 7500.00, 25.00],
    ["Product B", 310, 12600.00, 40.00],
    ["Product C", 190, 11400.00, 60.00],
    ["Product D", 125, 9500.00, 100.00],
    ["Product E", 135, 4475.00, 15.00]
]
elements.append(Paragraph("Productwise Performance", styles['Heading2']))
elements.append(create_clean_table(productwise_data, col_widths=[100, 100, 100, 150]))
elements.append(Spacer(1, 20))
product = """The report shows varied performance across products, with Product B leading in units sold, while Product A generates the highest total sales due to a higher unit price. Product C, despite lower sales volume, has a strong average price per unit, contributing significantly to revenue. Product D has the lowest sales, suggesting potential for improvement. Overall, focusing on boosting sales for underperforming products and leveraging the higher price points of others could optimize overall profitability."""
product_style = ParagraphStyle(name="Product Conclusion",fontSize=12,fontName="Helvetica",spaceAfter=20,alignment=TA_JUSTIFY,leading=15)
elements.append(Paragraph(product,product_style))

#Regionwise Table
regionwise_data = [
    ["Region", "Total Sales ", "Top-Selling Product"],
    ["Pune", "1,250.00", "Product A"],
    ["Mumbai", "1,200.00", "Product B"],
    ["Nashik", "2,500.00", "Product A"],
    ["Thane", "2,100.00", "Product C"],
    ["Nagpur", "2,400.00", "Product B"],
    ["Kolhapur", "3,000.00", "Product D"]
]
elements.append(Paragraph("Regionwise Performance", styles['Heading2']))
elements.append(create_clean_table(regionwise_data, col_widths=[150, 150, 150]))
elements.append(Spacer(1, 28))
region = """The region-wise sales data indicates strong regional performance, with Kolhapur leading in total sales at 3,000.00, driven by Product D. Pune and Nashik also show solid performance with Product A being the top-seller in both regions. Mumbai and Nagpur have notable sales figures with Product B topping the charts. Thane stands out with Product C as the top-seller, contributing to a total sales figure of 2,100.00."""
region_style = ParagraphStyle(name="Region Conclusion",fontSize=12,fontName="Helvetica",spaceAfter=20,alignment=TA_JUSTIFY,leading=15)
elements.append(Paragraph(region,region_style))


#Salespersonwise Table
salespersonwise_data = [
    ["Salesperson",	"Total Sales",	"Top-Selling Product"],
    ["Rajesh Deshmukh",	"1250.00",	"Product A"],
    ["Aishwarya Kulkarni",	"1200.00",	"Product B"],
    ["Priya Patil",	"1125.00",	"Product A"],
    ["Ramesh Joshi", "1200.00",	"Product C"],
    ["Anil Sharma",	"2400.00",	"Product B"]
]
elements.append(Paragraph("Salespersonwise Performance",styles['Heading2']))
elements.append(create_clean_table(salespersonwise_data,col_widths=[160,150,150]))
elements.append(Spacer(1,30))

#Discount & Sales Bar Graph
heading_style = ParagraphStyle(name="Heading",alignment=TA_CENTER,fontSize=18,fontName="Helvetica-Bold")
elements.append(Paragraph("Discount Applied over Total Sales",heading_style))
elements.append(Image("C:/Users/RUCHI/Codetech IT/discount_sales_bargraph.png", width=400, height=300))
elements.append(Spacer(1, 30))

#Conclusion:
conclusion = ["Total Revenue: The total sales amount from January 1, 2025, to January 12, 2025, is $36,850.00.",
              "Top-Performing Product Category: Product C generated the highest revenue with total sales of $12,600.00 across all regions.",
              "Highest Sales Region: The region Thane contributed significantly, with total revenue from various products amounting to $3,300.00.",
              "Lowest Sales Region: Raigad generated the lowest revenue, contributing $2,675.00, primarily through Products D and E.",
              "Best Day: January 10, 2025, recorded the highest daily revenue of $5,000.00, led by strong sales of Product D in the Raigad region."]
conclusion_style = ParagraphStyle(name="Conclusion Styling",fontSize=12,bulletIndent=15,spaceAfter=20,fontName="Helvetica")
for point in conclusion:
    elements.append(Paragraph(f"• {point}",conclusion_style))

#Pie Chart
heading2_style = ParagraphStyle(name="Heading2",alignment=TA_CENTER,fontSize=18,fontName="Helvetica-Bold",spaceAfter=20)
elements.append(Paragraph("Sales Trends by Product", heading2_style))
elements.append(Image("C:/Users/RUCHI/Codetech IT/sales_trends.png", width=400, height=300))
elements.append(Spacer(1, 20))

pie_chart = """The product-wise sales data highlights notable trends, with Product B leading overall sales at 27.4%, signifying its strong market performance. Product C follows closely with 23.7%, while Product D contributes 22.4% to the total sales, showcasing solid results. Product A accounts for 19.7%, indicating steady performance. However, Product E lags behind, representing only 6.9% of the total sales.

To maintain growth, strategic efforts should focus on enhancing the performance of Product E and leveraging the strengths of top-performing products like B and C to drive further sales growth."""
pie_chart_style = ParagraphStyle(name="Pie Chart Conclusion",fontSize=12,fontName="Helvetica",spaceAfter=18,alignment=TA_JUSTIFY,leading=15)
elements.append(Paragraph(pie_chart,pie_chart_style))

#Build PDF
doc.build(elements)
print(f"PDF saved at {pdf_path}")
