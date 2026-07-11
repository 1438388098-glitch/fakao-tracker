from flask import Flask, redirect, url_for
from config import Config
from database.db import db, init_db


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # Ensure data directory exists
    import os
    os.makedirs(os.path.join(app.root_path, "data"), exist_ok=True)

    init_db(app)

    # Register blueprints
    from routes.dashboard import dashboard_bp
    from routes.schedule import schedule_bp
    from routes.checkin import checkin_bp

    app.register_blueprint(dashboard_bp)
    app.register_blueprint(schedule_bp)
    app.register_blueprint(checkin_bp)

    @app.route("/")
    def index():
        return redirect(url_for("dashboard.home"))

    return app


if __name__ == "__main__":
    # Work around Windows hostname encoding bug
    import socket
    _orig_getfqdn = socket.getfqdn
    def _patched_getfqdn(name=""):
        try:
            return _orig_getfqdn(name)
        except UnicodeDecodeError:
            return name or "localhost"
    socket.getfqdn = _patched_getfqdn

    app = create_app()
    app.run(debug=True, host="127.0.0.1", port=5000)
