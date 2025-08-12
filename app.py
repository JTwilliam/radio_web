from flask import (Flask, render_template, request, redirect,
                   url_for, flash, session, send_file)
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import io, openpyxl

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///reg.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.secret_key = 'replace-me-in-production'

db = SQLAlchemy(app)

# ---------- 数据模型 ----------
class Reg(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(20))
    stu_id = db.Column(db.String(20), unique=False)
    major_class = db.Column(db.String(50))
    first_choice = db.Column(db.String(20))
    second_choice = db.Column(db.String(20))
    intro = db.Column(db.Text)
    time = db.Column(db.DateTime, default=datetime.utcnow)

class Config(db.Model):
    key = db.Column(db.String(20), primary_key=True)
    value = db.Column(db.String(10))

# ---------- 工具 ----------
def allow_register():
    c = Config.query.get('allow')
    return c.value == '1' if c else True

def allow_edit():
    c = Config.query.get('allow_edit')
    return c.value == '1' if c else True

# ---------- 普通用户 ----------
@app.route('/', methods=['GET', 'POST'])
def index():
    if not allow_register():
        return render_template('form.html', closed=True)

    if request.method == 'POST':
        name = request.form['name'].strip()
        stu_id = request.form['stu_id'].strip()
        exist = Reg.query.filter_by(stu_id=stu_id).first()
        if exist:
            if exist.name != name:
                return render_template('form.html',
                                       msg='该学号已被报名，若非本人报名，请联系管理人员:13360506019',
                                       closed=False, allow_edit=allow_edit())
            exist.first_choice = request.form['first']
            exist.second_choice = request.form['second']
            exist.intro = request.form['intro'][:200]
            db.session.commit()
            return render_template('form.html',
                                   msg='修改完成！', closed=False, already=True,
                                   name=name, stu_id=stu_id, allow_edit=allow_edit())
        else:
            db.session.add(Reg(
                name=name,
                stu_id=stu_id,
                major_class=request.form['major_class'],
                first_choice=request.form['first'],
                second_choice=request.form['second'],
                intro=request.form['intro'][:200]
            ))
            db.session.commit()
            return render_template('form.html',
                                   msg='报名成功！', closed=False, already=True,
                                   name=name, stu_id=stu_id, allow_edit=allow_edit())

    stu_id_cookie = session.get('stu_id')
    if stu_id_cookie:
        r = Reg.query.filter_by(stu_id=stu_id_cookie).first()
        if r:
            return render_template('form.html',
                                   closed=False, already=True,
                                   name=r.name, stu_id=r.stu_id,
                                   first=r.first_choice, second=r.second_choice,
                                   intro=r.intro, allow_edit=allow_edit())
    return render_template('form.html', closed=False, allow_edit=allow_edit())

@app.route('/edit', methods=['GET', 'POST'])
def edit():
    if not allow_register():
        return render_template('edit.html', closed=True)

    if not allow_edit():
        return render_template('edit.html', closed=False, edit_closed=True)

    if request.method == 'POST':
        name = request.form['name'].strip()
        stu_id = request.form['stu_id'].strip()
        rec = Reg.query.filter_by(stu_id=stu_id).first()
        if not rec:
            return render_template('edit.html', msg='未找到报名信息，请检查姓名和学号。')
        if rec.name != name:
            return render_template('edit.html', msg='姓名与学号不匹配，无法修改。')
        return render_template('edit_form.html', rec=rec)

    return render_template('edit.html')

# ---------- 管理员 ----------
@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST':
        if 'toggle' in request.form:
            c = Config.query.get('allow')
            c.value = '0' if c.value == '1' else '1'
            db.session.commit()
        if 'toggle_edit' in request.form:
            c = Config.query.get('allow_edit')
            c.value = '0' if c.value == '1' else '1'
            db.session.commit()
        return redirect(url_for('admin'))

    allow_reg = allow_register()
    allow_ed  = allow_edit()
    return render_template('admin.html', allow=allow_reg, allow_edit=allow_ed)

@app.route('/download')
def download():
    regs = Reg.query.order_by(Reg.first_choice, Reg.second_choice).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['姓名', '学号', '专业班级', '第一志愿', '第二志愿', '自我介绍', '报名时间'])
    for r in regs:
        ws.append([r.name, r.stu_id, r.major_class, r.first_choice,
                   r.second_choice, r.intro, r.time.strftime('%Y-%m-%d %H:%M')])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f'radio_club_{datetime.now():%Y%m%d}.xlsx')

# ---------- 预览 / 删除 ----------
@app.route('/preview')
def preview():
    records = Reg.query.all()
    return render_template('preview.html', records=records)

@app.route('/delete_one', methods=['POST'])
def delete_one():
    stu_id = request.form.get('stu_id')
    reg = Reg.query.filter_by(stu_id=stu_id).first_or_404()
    db.session.delete(reg)
    db.session.commit()
    flash('已删除一条记录')
    return redirect(url_for('preview'))

@app.route('/delete_all', methods=['POST'])
def delete_all():
    Reg.query.delete()
    db.session.commit()
    flash('已清空全部数据')
    return redirect(url_for('preview'))

# ---------- 初始化 ----------
with app.app_context():
    db.create_all()
    if not Config.query.get('allow'):
        db.session.add(Config(key='allow', value='1'))
    if not Config.query.get('allow_edit'):
        db.session.add(Config(key='allow_edit', value='1'))
    db.session.commit()

if __name__ == '__main__':
    app.run(debug=True)