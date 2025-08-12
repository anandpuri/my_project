from flask import Blueprint, request, jsonify, render_template, redirect, session, url_for
import os
import pandas as pd
import re

# ✅ Define blueprint
auth_bp = Blueprint('auth', __name__)

# ✅ Route: Fetch employee details (used in register.html via JavaScript)
@auth_bp.route('/get_user_details')
def get_user_details():
    emp_code = request.args.get('emp_code')
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    user_details_path = os.path.join(base_dir, 'data', 'user_details.xlsx')

    if not os.path.exists(user_details_path):
        return jsonify({'success': False, 'error': 'user_details.xlsx not found'})

    try:
        df = pd.read_excel(user_details_path)
        match = df[df['Employee Code'].astype(str) == str(emp_code)]

        if not match.empty:
            row = match.iloc[0]
            return jsonify({
                'success': True,
                'name': row['Name'],
                'department': row['Department'],
                'designation': row['Designation']
            })
        else:
            return jsonify({'success': False, 'error': 'Employee code not found.'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

# ✅ Route: Registration logic
@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    users_path = os.path.join(base_dir, 'data', 'users.xlsx')
    user_details_path = os.path.join(base_dir, 'data', 'user_details.xlsx')

    if request.method == 'POST':
        emp_code = request.form['emp_code']
        username = request.form['username']
        password = request.form['password']
        confirm = request.form['confirm']

        # ✅ Password rule validation
        if not re.match(r'^(?=.*[A-Z])(?=.*[a-zA-Z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$', password):
            error = "Password must be at least 8 characters with 1 capital letter, 1 special character, and numbers."
            return render_template('register.html', error=error)

        if password != confirm:
            error = "Passwords do not match."
            return render_template('register.html', error=error)

        # ✅ Load employee info from user_details.xlsx
        if not os.path.exists(user_details_path):
            return render_template('register.html', error="user_details.xlsx not found.")
        df_details = pd.read_excel(user_details_path)
        emp_match = df_details[df_details['Employee Code'].astype(str) == str(emp_code)]
        if emp_match.empty:
            return render_template('register.html', error="Employee code not found.")
        emp_row = emp_match.iloc[0]

        # ✅ Load or create users.xlsx
        if os.path.exists(users_path):
            df_users = pd.read_excel(users_path)
        else:
            df_users = pd.DataFrame(columns=['Username', 'Password', 'Employee Code', 'Name', 'Department', 'Designation'])

        # ✅ Check if employee already registered
        if not df_users[df_users['Employee Code'].astype(str) == str(emp_code)].empty:
            return render_template('register.html', error="This employee has already registered.")

        # ✅ Append new user
        new_user = {
            'Username': username,
            'Password': password,
            'Employee Code': emp_code,
            'Name': emp_row['Name'],
            'Department': emp_row['Department'],
            'Designation': emp_row['Designation']
        }
        df_users = df_users.append(new_user, ignore_index=True)
        df_users.to_excel(users_path, index=False)

        return redirect(url_for('auth.login'))

    return render_template('register.html')
