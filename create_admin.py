from app import app, db, User
from werkzeug.security import generate_password_hash

def create_admin():
    with app.app_context():
        existing_admin = User.query.filter_by(role='Admin').first()
        if existing_admin:
            print("❌ Admin sudah ada:", existing_admin.email)
            return

        name = "Daniel"
        phone = "081362109160"
        email = "daniel@uatas.id"
        password = "daniel766Hi"
        role = "Admin"

        hashed_pw = generate_password_hash(password)

        admin = User(
            name=name,
            phone=phone,
            email=email,
            password=hashed_pw,
            role=role
        )

        db.session.add(admin)
        db.session.commit()
        print("✅ Admin berhasil dibuat!")
        print(f"📧 Email: {email}")
        print(f"🔒 Password: {password} (gunakan untuk login pertama kali)")

if __name__ == '__main__':
    create_admin()