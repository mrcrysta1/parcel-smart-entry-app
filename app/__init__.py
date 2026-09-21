from flask import Flask, jsonify, request


def create_app():
    app = Flask(__name__, static_folder='static', template_folder='templates')
    # A workbook travels as base64 JSON, which inflates it by ~33%, so the
    # limit has to be well above the largest .xlsx we expect to handle.
    app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024
    app.config['JSON_SORT_KEYS'] = False

    from .routes import bp
    app.register_blueprint(bp)

    def wants_json():
        return request.path.startswith('/api/')

    @app.errorhandler(413)
    def too_large(e):
        msg = 'That file is too large. Please keep the workbook under about 45 MB.'
        return (jsonify({'error': msg}), 413) if wants_json() else (msg, 413)

    @app.errorhandler(404)
    def not_found(e):
        if wants_json():
            return jsonify({'error': 'Unknown endpoint.'}), 404
        return 'Not found', 404

    @app.errorhandler(Exception)
    def unhandled(e):
        code = getattr(e, 'code', 500)
        if not isinstance(code, int):
            code = 500
        if wants_json():
            app.logger.exception('API error')
            return jsonify({'error': str(getattr(e, 'description', None) or e)}), code
        raise e

    return app
