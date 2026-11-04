import streamlit as st
from docx import Document
from io import BytesIO
from datetime import date
import re, os, smtplib, ssl
from email.message import EmailMessage

st.set_page_config(page_title="Demande de projet | OV-Chimie-Métabolisme", page_icon="🧪", layout="wide")

st.markdown("""
<style>
.block-container {max-width: 1050px; padding-top: 2rem;}
.hero {padding: 1.4rem 1.6rem; border-radius: 16px; background: linear-gradient(120deg,#edf5fb,#f4f7f2); margin-bottom:1.2rem;}
.hero h1 {margin:0;color:#183b56;font-size:1.8rem;}
.hero p {margin:.4rem 0 0;color:#526575;}
div[data-testid="stForm"] {border:1px solid #dce5eb;border-radius:14px;padding:1.2rem;}
</style>
<div class="hero">
<h1>OV-Chimie-Métabolisme</h1>
<p>Formulaire de demande d'analyse — E-GES-01</p>
</div>
""", unsafe_allow_html=True)

st.write("Complétez les rubriques ci-dessous. À la dernière étape, vous pourrez télécharger le document Word et, si le service d'envoi est configuré, l'envoyer directement par e-mail.")
steps = ["Contacts", "Description du projet", "Échantillons", "Résultats & bibliographie", "Informations complémentaires", "Vérification & envoi"]
if "step" not in st.session_state: st.session_state.step=0
if "data" not in st.session_state: st.session_state.data={}
d=st.session_state.data

def field(key, label, **kwargs):
    d[key]=st.text_input(label, value=d.get(key,""), key="f_"+key, **kwargs)
def area(key,label,**kwargs):
    d[key]=st.text_area(label, value=d.get(key,""), height=kwargs.pop("height",150), key="f_"+key, **kwargs)

st.progress((st.session_state.step+1)/len(steps), text=f"Étape {st.session_state.step+1}/{len(steps)} — {steps[st.session_state.step]}")
with st.container(border=True):
    if st.session_state.step==0:
        st.subheader("1. Informations générales")
        c1,c2=st.columns(2)
        with c1:
            field("project_manager","Responsable du projet *")
            field("group","Groupe / laboratoire")
            d["date"]=st.date_input("Date", value=date.fromisoformat(d["date"]) if d.get("date") else date.today()).isoformat()
            field("manager_phone","Téléphone")
            field("manager_email","E-mail du responsable")
        with c2:
            field("admin_staff","Personnel administratif")
            field("admin_phone","Téléphone du personnel administratif")
            field("admin_email","E-mail du personnel administratif")
        st.caption("* Champ recommandé.")
    elif st.session_state.step==1:
        st.subheader("2. Description du projet")
        st.caption("Questions scientifiques et techniques à traiter, molécules à analyser.")
        area("project_description","Description scientifique et technique",height=230)
    elif st.session_state.step==2:
        st.subheader("3. Nature et nombre des échantillons")
        st.caption("Indiquez notamment les références du procédé analytique, les solvants et les formules chimiques si disponibles.")
        area("samples","Nature et nombre des échantillons / informations analytiques",height=220)
    elif st.session_state.step==3:
        st.subheader("4. Résultats préliminaires et bibliographie")
        area("preliminary","Résultats déjà obtenus et éléments utiles à l'évaluation du projet",height=160)
        area("bibliography","Résumé bibliographique et références des articles pertinents",height=160)
    elif st.session_state.step==4:
        st.subheader("5. Informations complémentaires")
        field("project_name","Nom du projet *")
        d["funding"]=st.radio("Financement (uniquement pour les chercheurs IJPB)",["YES","NO"], horizontal=True,index=["YES","NO"].index(d.get("funding","NO")))
        d["guidance"]=st.radio("Souhaitez-vous un accompagnement scientifique ?",["YES","NO"],horizontal=True,index=["YES","NO"].index(d.get("guidance","NO")))
        st.info("La plateforme peut vous mettre en relation avec des scientifiques de l'IJPB pour aider à interpréter les résultats. Une co-signature / co-autorat pourra alors être requise.")
    else:
        st.subheader("6. Vérification, génération et envoi")
        st.write("Vérifiez les informations avant de générer le document.")
        summary = [
            ("Responsable",d.get("project_manager","")),("Groupe",d.get("group","")),
            ("Date",d.get("date","")),("E-mail",d.get("manager_email","")),
            ("Projet",d.get("project_name","")),("Financement IJPB",d.get("funding","NO")),
            ("Accompagnement scientifique",d.get("guidance","NO"))
        ]
        for k,v in summary: st.markdown(f"**{k} :** {v or '—'}")
        with st.expander("Afficher toutes les réponses"):
            for k,v in d.items():
                st.markdown(f"**{k.replace('_',' ').capitalize()} :** {v}")
        st.divider()
        recipient=st.text_input("Envoyer le document à cette adresse e-mail",value="ov-chimie@inrae.fr")
        subject=st.text_input("Objet du message",value=f"Demande de projet – {d.get('project_name','E-GES-01')}")
        sender_name=st.text_input("Nom de l'expéditeur (facultatif)",value=d.get("project_manager",""))
        st.caption("L'envoi automatique nécessite la configuration d'un compte SMTP dans les secrets de l'application. Sans cette configuration, le téléchargement Word reste disponible.")

def set_cell(cell, label, value):
    # Retain original cell content and append the submitted value below it.
    if not value: return
    p=cell.add_paragraph()
    p.paragraph_format.space_before=Pt(3)
    r=p.add_run(value)
    r.font.size=Pt(9)

def build_doc(data):
    doc=Document(os.path.join(os.path.dirname(__file__),"template_E-GES-01.docx"))
    # Contact fields: template table 1
    t=doc.tables[1]
    set_cell(t.cell(0,0),"",data.get("project_manager",""))
    set_cell(t.cell(0,2),"",data.get("group",""))
    set_cell(t.cell(1,0),"",data.get("date",""))
    set_cell(t.cell(1,1),"",data.get("manager_phone",""))
    set_cell(t.cell(1,3),"",data.get("manager_email",""))
    set_cell(t.cell(2,0),"",data.get("admin_staff",""))
    set_cell(t.cell(2,1),"",data.get("admin_phone",""))
    set_cell(t.cell(2,3),"",data.get("admin_email",""))
    set_cell(t.cell(3,0),"",data.get("project_description",""))
    set_cell(t.cell(4,0),"",data.get("samples",""))
    set_cell(t.cell(5,0),"",data.get("preliminary",""))
    if data.get("bibliography"):
        set_cell(t.cell(5,0),"","Bibliographie : "+data["bibliography"])
    set_cell(t.cell(6,0),"",data.get("project_name",""))
    # Keep original YES/NO fields; append selected values explicitly.
    set_cell(t.cell(6,0),"","Financement IJPB : "+data.get("funding","NO"))
    set_cell(t.cell(7,0),"","Accompagnement scientifique : "+data.get("guidance","NO"))
    # Update header reference text where present
    ref=f"E-GES-01_{date.today().strftime('%d%m%Y')}_{re.sub(r'[^A-Za-z0-9]+','_',data.get('project_manager','NOM')).strip('_').upper() or 'NOM'}"
    for table in doc.tables[:1]:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "E-GES-01_AAMMJJ_NOM" in p.text:
                        for run in p.runs: run.text=run.text.replace("E-GES-01_AAMMJJ_NOM",ref)
    out=BytesIO(); doc.save(out); out.seek(0)
    return out.getvalue(), ref+".docx"

if st.session_state.step < len(steps)-1:
    b1,b2=st.columns([1,1])
    with b1:
        if st.button("← Précédent",disabled=st.session_state.step==0,use_container_width=True):
            st.session_state.step-=1; st.rerun()
    with b2:
        if st.button("Suivant →",use_container_width=True):
            st.session_state.step+=1; st.rerun()
else:
    if st.button("← Modifier les réponses"):
        st.session_state.step-=1; st.rerun()
    if not d.get("project_manager") or not d.get("project_name"):
        st.warning("Renseignez au minimum le responsable et le nom du projet avant la génération.")
    else:
        content, filename=build_doc(d)
        st.download_button("📄 Télécharger le document Word",data=content,file_name=filename,mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
        if st.button("✉️ Envoyer le document par e-mail",use_container_width=True):
            try:
                cfg=st.secrets["smtp"]
                msg=EmailMessage()
                msg["Subject"]=subject
                msg["From"]=cfg["sender"]
                msg["To"]=recipient
                msg.set_content(f"Bonjour,\\n\\nVeuillez trouver ci-joint la demande de projet E-GES-01.\\n\\nCordialement,\\n{sender_name}")
                msg.add_attachment(content,maintype="application",subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",filename=filename)
                with smtplib.SMTP(cfg["host"],int(cfg.get("port",587))) as server:
                    server.starttls(context=ssl.create_default_context())
                    server.login(cfg["username"],cfg["password"])
                    server.send_message(msg)
                st.success(f"Document envoyé à {recipient}.")
            except Exception as e:
                st.error("Envoi non configuré ou impossible. Le document reste téléchargeable. Détail : "+str(e))
