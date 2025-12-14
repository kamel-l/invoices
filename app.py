from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import os
from flask_migrate import Migrate

app = Flask(__name__)
app.config['SECRET_KEY'] = 'votre_clé_secrète_ici'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///inventaire.db'
db = SQLAlchemy(app)
migrate = Migrate(app, db)

class Produit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nom = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text, nullable=True)
    quantite = db.Column(db.Integer, nullable=False)
    prix = db.Column(db.Float, nullable=False)
    date_ajout = db.Column(db.DateTime, default=datetime.utcnow)



@app.route('/')
def index():
    produits = Produit.query.all()
    return render_template('index.html', produits=produits)

@app.route('/ajouter', methods=['GET', 'POST'])
def ajouter_produit():
    if request.method == 'POST':
        nom = request.form['nom']
        description = request.form['description']
        quantite = int(request.form['quantite'])
        prix = float(request.form['prix'])

        nouveau_produit = Produit(nom=nom, description=description, quantite=quantite, prix=prix)
        db.session.add(nouveau_produit)
        db.session.commit()
        flash('Produit ajouté avec succès!', 'success')
        return redirect(url_for('index'))
    return render_template('ajouter.html')

@app.route('/modifier/<int:id>', methods=['GET', 'POST'])
def modifier_produit(id):
    produit = Produit.query.get_or_404(id)
    if request.method == 'POST':
        produit.nom = request.form['nom']
        produit.description = request.form['description']
        produit.quantite = int(request.form['quantite'])
        produit.prix = float(request.form['prix'])
        db.session.commit()
        flash('Produit modifié avec succès!', 'success')
        return redirect(url_for('index'))
    return render_template('modifier.html', produit=produit)

@app.route('/supprimer/<int:id>')
def supprimer_produit(id):
    produit = Produit.query.get_or_404(id)
    db.session.delete(produit)
    db.session.commit()
    flash('Produit supprimé avec succès!', 'success')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
