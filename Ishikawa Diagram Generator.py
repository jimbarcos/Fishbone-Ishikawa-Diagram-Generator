import tkinter as tk
import tkinter.messagebox as messagebox
import pandas as pd
from tkinter import filedialog

def process_excel_file():
  # Create a Tkinter window
  root = tk.Tk()
  root.withdraw()

  # Prompt the user to select an Excel file
  file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

  # Read the Excel file into a pandas DataFrame
  try:
    df = pd.read_excel(file_path)
  except Exception:
    print("error reading file")
    exit()

  # Print the DataFrame/Column Based
  print(df)

  # Print the DataFrame
  filtered_data = {}
  for column in df.columns:
    filtered_data[column] = df[column][df[column].apply(lambda x: isinstance(x, str) and any(c.isalpha() for c in x))]

  # Print or use the filtered data as needed
  for column, data in filtered_data.items():
    print(f"{column}: {data}")

  combined_data = pd.concat(list(filtered_data.values()), ignore_index=False)

  try:
    column_1_data = filtered_data['Unnamed: 1']
  except Exception:
    print("Column 1 not found in file")

  try:
    column_2_data = filtered_data['Unnamed: 2']
  except Exception:
    print("Column 2 not found in file")

  try:
    column_3_data = filtered_data['Unnamed: 3']
  except Exception:
    print("Column 3 not found in file")

  # Passed the parameters to the function
  try:
    root.destroy()
    draw_fishbone_diagram(combined_data, column_1_data, column_2_data, column_3_data)
  except Exception:
    print("Invalid File Format. Please select a valid file.")
    exit()

def draw_fishbone_diagram(cdata, column_1_data, column_2_data, column_3_data):
  cdata = cdata.sort_index()
  column_1_data = column_1_data.sort_index()
  column_2_data = column_2_data.sort_index()
  column_3_data = column_3_data.sort_index()

  # Print the DataFrame/Row Based
  print(cdata)

  # Create the main window
  window = tk.Tk()
  window.title("Fishbone Diagram")

  # Calculate the center position of the screen
  screen_width = window.winfo_screenwidth()
  screen_height = window.winfo_screenheight()
  window_width = 1000
  window_height = 600
  wx = (screen_width - window_width) // 2
  wy = (screen_height - window_height) // 2

  # Set the window position
  window.geometry(f"{window_width}x{window_height}+{wx}+{wy}")

  # Disable window resizing
  window.resizable(False, False)

  # Set the foreground color for text
  outlineBone, branch, subComponents, components, bgColor, title = "#6dcff6", "#fffacc", "#df5b56", "#e2bd06", "#1E1E1E", "#F5F5F5"
  rectangleColor, upperTitle, branchBG = "#021f13", "#0de3ff", "#4c866b"

  # Create the canvas to draw the diagram
  canvas = tk.Canvas(window, width=1000, height=600, bg=bgColor, highlightthickness=5)
  canvas.pack()

  # Draw the main horizontal line
  canvas.create_line(50, 300, 750, 300, width=2, fill=outlineBone)
  # Draw the branches
  canvas.create_line(160, 100, 270, 300, width=2, fill=outlineBone)  # Branch 1
  canvas.create_line(380, 100, 500, 300, width=2, fill=outlineBone)  # Branch 2
  canvas.create_line(620, 100, 750, 300, width=2, fill=outlineBone)  # Branch 3
  canvas.create_line(160, 500, 270, 300, width=2, fill=outlineBone)  # Branch 4
  canvas.create_line(380, 500, 500, 300, width=2, fill=outlineBone)  # Branch 5
  canvas.create_line(620, 500, 750, 300, width=2, fill=outlineBone)  # Branch 6
  # Draw the title head
  canvas.create_rectangle(750, 200, 950, 400, width=2, outline=outlineBone)
  
  x_adjust, y_adjust = 100, 450
  # Add labels for the Legend
  canvas.create_text(812+x_adjust,20+y_adjust, text="Color Legend: ", fill=title, font=("Tahoma", 11, "bold"))
  canvas.create_text(783+x_adjust,40+y_adjust, text="• Title", fill=title, font=("Tahoma", 11, "bold"))
  canvas.create_text(792+x_adjust,60+y_adjust, text="• Branch", fill=branch, font=("Tahoma", 11, "bold"))
  canvas.create_text(811+x_adjust,80+y_adjust, text="• Components", fill=components, font=("Tahoma", 11, "bold"))
  canvas.create_text(820+x_adjust,100+y_adjust, text="• Subomponents", fill=subComponents, font=("Tahoma", 11, "bold"))
  canvas.create_text(793+x_adjust,120+y_adjust, text="• Outline", fill=outlineBone, font=("Tahoma", 11, "bold"))
  #Add title and subtitle
  canvas.create_rectangle(0, 0, 1001, 60, fill=rectangleColor)
  canvas.create_text(500,20, text="> Cause & Effect Fishbone Diagram <", fill=upperTitle, font=("Tahoma", 18, "bold"))
  canvas.create_text(500,50, text=f"title: \"{cdata[0]}\"", fill=title, font=("Tahoma", 10, "bold"))

  # Add labels for the branches
  num_branches = len(cdata)
  try:
    branch_labels = [f"{cdata[i+1]}" for i in range(num_branches-1)]
  except Exception:
    print("error reading file")
    exit()

  # Add foundation display for the branches
  canvas.create_rectangle(90, 70, 230, 90, fill=branchBG)
  canvas.create_rectangle(310, 70, 450, 90, fill=branchBG)
  canvas.create_rectangle(550, 70, 690, 90, fill=branchBG)
  canvas.create_rectangle(90, 505, 230, 525, fill=branchBG)
  canvas.create_rectangle(310, 505, 450, 525, fill=branchBG)
  canvas.create_rectangle(550, 505, 690, 525, fill=branchBG)
  
  counter, seed, seed1, seed2, seed3, seed4, seed5,col3 = 0, 0, 0, 0, 0, 0, 0, 0
  for i, label in enumerate(branch_labels):
    x = 150 
    y = 80 + i * 30

    # If the label is in the first column, then it is a branch
    if i in column_1_data.index-1:
      if (counter == 0):
        canvas.create_text(x+10, 80, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 1):
        canvas.create_text(x+230, 80, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 2):
        canvas.create_text(x+470, 80, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 3):
        canvas.create_text(x+10,  515, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 4):
        canvas.create_text(x+230, 515, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 5):
        canvas.create_text(x+470, 515, text=label, fill=branch, font=("Tahoma", 11, "bold"))
        counter += 1
      elif (counter == 6):
        counter += 1

    # If the label is in the second column, then it is a component
    if i in column_2_data.index-1:
      if (counter == 1):
        canvas.create_text(80, y, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(130, y, 130+30, y, width=2, fill=components, arrow=tk.LAST)
        seed += 1
      elif (counter == 2):
        y = 80 + (i-seed-1) * 30  # Reset the value of y
        canvas.create_text(300, y, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(350, y, 350+30, y, width=2, fill=components, arrow=tk.LAST)
        seed1 += 1
      elif (counter == 3):
        y = 80 + (i-(seed+seed1)-3) * 30  # Reset the value of y
        canvas.create_text(540, y+30, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(590, y+30, 590+30, y+30, width=2, fill=components, arrow=tk.LAST)
        seed2 += 1
        col3 += 1
      elif (counter == 4):
        y = 80 + (i-(seed+seed1+seed2)-3) * 30  # Reset the value of y
        canvas.create_text(80, y+140, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(130, y+140, 130+30, y+140, width=2, fill=components, arrow=tk.LAST)
        seed3 += 1
      elif (counter == 5):
        y = 80 + (i-(seed+seed1+seed2+seed3)-4) * 30  # Reset the value of y
        canvas.create_text(300, y+140, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(350, y+140, 350+30, y+140, width=2, fill=components, arrow=tk.LAST)
        seed4 += 1
      elif (counter == 6):
        y = 80 + (i-(seed+seed1+seed2+seed3+seed4)-5) * 30  # Reset the value of y
        canvas.create_text(540, y+140, text=label, fill=components, font=("Tahoma", 8, "bold"))
        canvas.create_line(590, y+140, 590+30, y+140, width=2, fill=components, arrow=tk.LAST)
        seed5 += 1
      elif (counter == 7):
        y = y

    # If the label is in the third column, then it is a subcomponent
    if i in column_3_data.index-1:
      if (counter == 1):
        canvas.create_text(80, y, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(130, y, 130, y-20, width=2, fill=subComponents, arrow=tk.LAST)
        seed += 1
      elif (counter == 2):
        y = 80 + (i-seed-1) * 30  # Reset the value of y
        canvas.create_text(300, y, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(350, y, 350, y-20, width=2, fill=subComponents, arrow=tk.LAST)
        seed1 += 1
      elif (counter == 3):
        y = 80 + (i-(seed+seed1)-3) * 30  # Reset the value of y
        canvas.create_text(540, y+30, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(590, y+30, 590, y+10, width=2, fill=subComponents, arrow=tk.LAST)
        col3 += 1
        seed2
      elif (counter == 4):
        y = 80 + (i-(seed+seed1+seed2)-3) * 30  # Reset the value of y
        canvas.create_text(80, y+140, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(130, y+140, 130, y+120, width=2, fill=subComponents, arrow=tk.LAST)
        seed3 += 1
      elif (counter == 5):
        y = 80 + (i-(seed+seed1+seed2+seed3)-4) * 30  # Reset the value of y
        canvas.create_text(300, y+140, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(350, y+140, 350, y+120, width=2, fill=subComponents, arrow=tk.LAST)
        seed4 += 1
      elif (counter == 6):
        y = 80 + (i-(seed+seed1+seed2+seed3+seed4)-5) * 30  # Reset the value of y
        canvas.create_text(540, y+140, text=label, fill=subComponents, font=("Tahoma", 8, "bold"))
        canvas.create_line(590, y+140, 590, y+120, width=2, fill=subComponents, arrow=tk.LAST)
        seed5 += 1
      elif (counter == 7):
        y = y
        
  textTitle = cdata[0]
  if len(textTitle) > 20:
    textTitle = textTitle[:20] + "\n" + textTitle[20:]
  if len(textTitle) > 40:
    textTitle = textTitle[:40] + "\n" + textTitle[40:]
  if len(textTitle) > 60:
    textTitle = textTitle[:60] + "\n" + textTitle[60:]
  if len(textTitle) > 80:
    textTitle = textTitle[:80] + "\n" + textTitle[80:]
  if len(textTitle) > 100:
    textTitle = textTitle[:100] + "\n" + textTitle[100:]
  canvas.create_text(845, 280, text=textTitle, fill=title, font=("Tahoma", 11, "bold"))

  if seed1 > 7 or col3 > 7 or seed3 > 7 or seed4 > 7 or seed5 > 7:
    messagebox.showwarning("Notice - Too Many Items", "Please note that the diagram may not be accurate if the number of components/subcomponents are more than 7 in one or more branches.")
  # Start the main event loop
 
  window.mainloop()

def main():
  process_excel_file()

if __name__ == "__main__":
  main()