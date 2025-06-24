import customtkinter as ctk
import requests
import json
import time
import os
from datetime import datetime
from threading import Thread
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np

# Cargar variables de entorno
from dotenv import load_dotenv
load_dotenv()

# --- CONFIGURACIÓN DE LA API DE INSTAGRAM ---
INSTAGRAM_APP_ID = os.environ.get('INSTAGRAM_APP_ID')
INSTAGRAM_COOKIE = os.environ.get('INSTAGRAM_COOKIE')

if not INSTAGRAM_APP_ID or not INSTAGRAM_COOKIE:
    print("Error: INSTAGRAM_APP_ID o INSTAGRAM_COOKIE no encontrados en el archivo .env.")
    print("Por favor, crea un archivo .env en la misma carpeta y añade tus credenciales.")
    print("Consulta la documentación para saber cómo obtener estas credenciales.")
    exit()

def get_headers():
    csrf_token = ""
    if INSTAGRAM_COOKIE:
        for part in INSTAGRAM_COOKIE.split(';'):
            if 'csrftoken=' in part:
                csrf_token = part.split('csrftoken=')[1].strip()
                break

    if not csrf_token:
        print("Advertencia: No se pudo extraer csrftoken de la cookie.")

    return {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:139.0) Gecko/20100101 Firefox/139.0',
        'x-ig-app-id': INSTAGRAM_APP_ID,
        'Cookie': INSTAGRAM_COOKIE,
        'X-CSRFToken': csrf_token,
    }

def fetch_user_profile(username):
    url = f"https://i.instagram.com/api/v1/users/web_profile_info/?username={username}"
    try:
        response = requests.get(url, headers=get_headers(), timeout=10)
        response.raise_for_status()
        data = response.json()

        if response.status_code == 401:
            return {"error": "Unauthorized", "message": "Revisa tus cookies o autenticación."}
        if response.status_code == 429:
            return {"error": "Rate-limited", "message": "Demasiadas peticiones. Espera un tiempo."}
        if response.status_code == 404:
            return {"error": "Not Found", "message": "Usuario no encontrado"}

        if 'data' not in data or 'user' not in data['data']:
            return {"error": "Invalid Response", "message": "Estructura de respuesta inesperada."}

        return data['data']['user']
    except requests.exceptions.Timeout:
        return {"error": "Timeout", "message": "La petición excedió el tiempo límite."}
    except requests.exceptions.RequestException as e:
        return {"error": "Request Error", "message": str(e)}
    except json.JSONDecodeError as e:
        return {"error": "JSON Parse Error", "message": str(e)}
    except Exception as e:
        return {"error": "Unknown Error", "message": str(e)}

def fetch_user_media(user_data, username, posts_count):
    if not user_data or "error" in user_data:
        return user_data

    user_id = user_data['id']
    all_timeline_media = []
    next_max_id = None
    
    TARGET_POSTS_COUNT = posts_count
    
    while len(all_timeline_media) < TARGET_POSTS_COUNT:
        remaining_posts = TARGET_POSTS_COUNT - len(all_timeline_media)
        request_count = min(remaining_posts, 50) 

        url = f"https://i.instagram.com/api/v1/feed/user/{user_id}/?count={request_count}"
        if next_max_id:
            url += f"&max_id={next_max_id}"

        try:
            response = requests.get(url, headers=get_headers(), timeout=10)
            
            if response.status_code == 401:
                return {"error": "Unauthorized", "message": "Revisa tus cookies o autenticación."}
            if response.status_code == 429:
                return {"error": "Rate-limited", "message": "Demasiadas peticiones. Espera un tiempo."}
            if response.status_code == 404: 
                return {"error": "Not Found", "message": f"No se encontraron medios para el usuario {username}."}
            
            response.raise_for_status() 
            data = response.json()
            
            current_items = data.get('items', [])
            all_timeline_media.extend(current_items)

            next_max_id = data.get('next_max_id') or data.get('next_max_id_v2')
            
            if not data.get('more_available') or not next_max_id or not current_items:
                break 

            time.sleep(1) 

        except requests.exceptions.Timeout:
            return {"error": "Timeout", "message": "La petición de medios excedió el tiempo límite."}
        except requests.exceptions.RequestException as e:
            return {"error": "Request Error", "message": f"Error al obtener medios para {username}: {e}"}
        except json.JSONDecodeError as e:
            return {"error": "JSON Parse Error", "message": f"Error al parsear JSON para {username}: {e}. Respuesta: {response.text[:200]}..."}
        except Exception as e:
            return {"error": "Unknown Error", "message": f"Error desconocido al obtener medios para {username}: {e}"}

    timeline_media = all_timeline_media[:TARGET_POSTS_COUNT]

    total_likes = 0
    total_comments = 0
    media_count_for_averages = 0 
    
    recent_media_data = []

    for item in timeline_media:
        if not item:
            continue

        caption = item.get('caption', {}).get('text', 'No caption')
        media_url = item.get('image_versions2', {}).get('candidates', [{}])[0].get('url', 'No media URL')
        like_count = item.get('like_count', 0)
        comment_count = item.get('comment_count', 0)
        
        timestamp = item.get('taken_at', 0) * 1000 

        if like_count is not None and comment_count is not None:
            total_likes += like_count
            total_comments += comment_count
            media_count_for_averages += 1 
        
        recent_media_data.append({
            "PostId": item.get('pk'),
            "Type": 'Foto' if item.get('media_type') == 1 else ('Video' if item.get('media_type') == 2 else 'Otro'),
            "Likes": like_count,
            "Comments": comment_count,
            "Caption": caption[:70] + '...' if len(caption) > 70 else caption,
            "PostDate": datetime.fromtimestamp(timestamp / 1000).strftime('%d/%m/%Y') if timestamp else 'N/A',
            "MediaUrl": media_url,
            "EsCarrusel": 'Sí' if item.get('carousel_media') else 'No'
        })

    average_likes = round(total_likes / media_count_for_averages, 2) if media_count_for_averages > 0 else 'No Disponible'
    average_comments = round(total_comments / media_count_for_averages, 2) if media_count_for_averages > 0 else 'No Disponible'

    followers_count = user_data.get('edge_followed_by', {}).get('count', 0)
    
    engagement_rate = 'No Disponible'
    if media_count_for_averages > 0 and followers_count > 0:
        engagement_rate = round(((average_likes + average_comments) / followers_count) * 100, 2)
        
    is_business_account = 'Sí' if user_data.get('is_business_account') else 'No'
    has_highlight_reels = 'Sí' if user_data.get('highlight_reel_count', 0) > 0 else 'No'
    external_url = user_data.get('external_url', 'No disponible')
    public_email = user_data.get('public_email', 'No disponible')
    public_phone_number = user_data.get('public_phone_number', 'No disponible')
    has_public_contact = 'Sí' if (public_email != 'No disponible' or public_phone_number != 'No disponible') else 'No'
    biography_has_links = 'Sí' if user_data.get('biography') and ('http://' in user_data['biography'] or 'https://' in user_data['biography']) else 'No'
    is_private = 'Sí' if user_data.get('is_private') else 'No'

    account_info = {
        "Nombre de usuario": user_data.get('username'),
        "Nombre completo": user_data.get('full_name', 'No Disponible'),
        "Biografía": user_data.get('biography', 'No Disponible'),
        "País": user_data.get('country_block', 'No Disponible'),
        "URL Perfil": f"https://www.instagram.com/{user_data.get('username')}",
        "Categoría": user_data.get('category_name', 'No Disponible'),
        "cantidad seguidores": followers_count,
        "cantidad seguidos": user_data.get('edge_follow', {}).get('count', 'No Disponible'),
        "cantidad de publicaciones": user_data.get('edge_owner_to_timeline_media', {}).get('count', 'No Disponible'),
        "Está verificado ✅": 'Sí' if user_data.get('is_verified') else 'No',
        "¿Es una cuenta profesional?": is_business_account,
        "Tiene Historias Destacadas": has_highlight_reels,
        "URL Externa (Bio)": external_url,
        "Email Público": public_email,
        "Teléfono Público": public_phone_number,
        "Tiene Contacto Público": has_public_contact,
        "Biografía con Links": biography_has_links,
        "Es Cuenta Privada": is_private,
        # Guardar la cantidad de posts usados para el promedio en los datos para mostrarlo en la GUI
        "Posts para promedio": media_count_for_averages,
        "Me gusta promedio 👍": average_likes,
        "Comentarios promedio 💬": average_comments,
        "Tasa de interacción 📊": f"{engagement_rate}%" if engagement_rate != 'No Disponible' else engagement_rate,
        "Últimos X Posts": recent_media_data
    }
    return account_info

def scrape_instagram_profiles(usernames_list, posts_to_fetch, callback):
    """Función principal de scraping que reporta el progreso a la GUI."""
    all_accounts_data = []
    
    for i, username in enumerate(usernames_list):
        clean_username = username.strip().replace('@', '')
        
        callback("log", f"[{i+1}/{len(usernames_list)}] Analizando perfil: {clean_username}...")
        
        user_profile = fetch_user_profile(clean_username)
        
        account_info = {}
        if user_profile and "error" not in user_profile:
            account_info = fetch_user_media(user_profile, clean_username, posts_to_fetch)
            if account_info and "error" not in account_info:
                all_accounts_data.append(account_info)
                callback("log", f"Datos completos obtenidos para {clean_username}.")
            else:
                placeholder_info = {
                    "Nombre de usuario": clean_username,
                    "Nombre completo": user_profile.get('full_name', 'No Disponible'),
                    "Biografía": user_profile.get('biography', 'No Disponible'),
                    "País": user_profile.get('country_block', 'No Disponible'),
                    "URL Perfil": f"https://www.instagram.com/{clean_username}",
                    "Categoría": user_profile.get('category_name', 'No Disponible'),
                    "cantidad seguidores": user_profile.get('edge_followed_by', {}).get('count', 'No Disponible'),
                    "cantidad seguidos": user_profile.get('edge_follow', {}).get('count', 'No Disponible'),
                    "cantidad de publicaciones": user_profile.get('edge_owner_to_timeline_media', {}).get('count', 'No Disponible'),
                    "Está verificado ✅": 'Sí' if user_profile.get('is_verified') else 'No',
                    "¿Es una cuenta profesional?": 'Sí' if user_profile.get('is_business_account') else 'No',
                    "Tiene Historias Destacadas": 'Sí' if user_profile.get('highlight_reel_count', 0) > 0 else 'No',
                    "URL Externa (Bio)": user_profile.get('external_url', 'No disponible'),
                    "Email Público": user_profile.get('public_email', 'No disponible'),
                    "Teléfono Público": user_profile.get('public_phone_number', 'No disponible'),
                    "Tiene Contacto Público": 'Sí' if (user_profile.get('public_email') or user_profile.get('public_phone_number')) else 'No',
                    "Biografía con Links": 'Sí' if user_profile.get('biography') and ('http://' in user_profile['biography'] or 'https://' in user_profile['biography']) else 'No',
                    "Es Cuenta Privada": 'Sí' if user_profile.get('is_private') else 'No',
                    "Posts para promedio": 0, # No se pudieron obtener posts, así que 0
                    "Me gusta promedio 👍": 'No Disponible',
                    "Comentarios promedio 💬": 'No Disponible',
                    "Tasa de interacción 📊": 'No Disponible',
                    "Últimos X Posts": [],
                    "Error al obtener medios": account_info.get("message", "Error desconocido al obtener medios") if "error" in account_info else "N/A"
                }
                all_accounts_data.append(placeholder_info)
                callback("log", f"Error o datos incompletos para {clean_username}: {account_info.get('message', 'N/A')}. Añadiendo datos parciales.")
        else:
            error_message = user_profile.get("message", "Error desconocido") if user_profile and "error" in user_profile else "No Disponible"
            
            if error_message == "Usuario no encontrado":
                callback("log", f"El usuario ingresado '{clean_username}' no existe.")
            else:
                callback("log", f"Error al obtener perfil para {clean_username}: {error_message}. Añadiendo datos no disponibles.")
            
            placeholder_info = {
                "Nombre de usuario": clean_username,
                "Nombre completo": 'No Disponible', "Biografía": 'No Disponible', "País": 'No Disponible',
                "URL Perfil": f"https://www.instagram.com/{clean_username}", "Categoría": 'No Disponible',
                "cantidad seguidores": 'No Disponible', "cantidad seguidos": 'No Disponible', "cantidad de publicaciones": 'No Disponible',
                "Está verificado ✅": 'No Disponible', "¿Es una cuenta profesional?": 'No Disponible', "Tiene Historias Destacadas": 'No Disponible',
                "URL Externa (Bio)": 'No Disponible', "Email Público": 'No Disponible', "Teléfono Público": 'No Disponible',
                "Tiene Contacto Público": 'No Disponible', "Biografía con Links": 'No Disponible', "Es Cuenta Privada": 'No Disponible',
                "Posts para promedio": 0, # No hay posts disponibles
                "Me gusta promedio 👍": 'No Disponible', "Comentarios promedio 💬": 'No Disponible',
                "Tasa de interacción 📊": 'No Disponible', "Últimos X Posts": [],
                "Error al obtener perfil": error_message
            }
            all_accounts_data.append(placeholder_info)
            
        time.sleep(2) 

    callback("results", all_accounts_data)
    callback("log", "Scraping completado.")
    return all_accounts_data

def export_to_excel_with_pivot_and_charts(data, filepath):
    if not data:
        return False

    try:
        df_accounts_raw = pd.DataFrame(data)
        df_accounts = df_accounts_raw.copy()

        for col in ["cantidad seguidores", "cantidad seguidos", "cantidad de publicaciones", 
                    "Me gusta promedio 👍", "Comentarios promedio 💬", "Posts para promedio"]: # Añadido Posts para promedio
            if col in df_accounts.columns:
                df_accounts[col] = df_accounts[col].replace('No Disponible', np.nan)
                df_accounts[col] = pd.to_numeric(df_accounts[col], errors='coerce').fillna(0)
        
        if "Tasa de interacción 📊" in df_accounts.columns:
            df_accounts["Tasa de interacción 📊"] = df_accounts["Tasa de interacción 📊"].astype(str).str.replace('%', '').replace('No Disponible', np.nan)
            df_accounts["Tasa de interacción 📊"] = pd.to_numeric(df_accounts["Tasa de interacción 📊"], errors='coerce').fillna(0)


        flat_posts_data = []
        for account in data:
            username = account.get("Nombre de usuario", "N/A")
            for post in account.get("Últimos X Posts", []):
                post_copy = post.copy()
                post_copy["Nombre de usuario"] = username
                flat_posts_data.append(post_copy)
        
        df_posts = pd.DataFrame(flat_posts_data)

        for col in ["Likes", "Comments"]:
            if col in df_posts.columns:
                df_posts[col] = pd.to_numeric(df_posts[col], errors='coerce').fillna(0)

        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            df_accounts.to_excel(writer, sheet_name='Datos Cuentas', index=False)
            
            df_posts.to_excel(writer, sheet_name='Datos Posts', index=False)

            workbook = writer.book
            worksheet_analysis = workbook.add_worksheet('Análisis Cuentas')

            pivot_values = ["cantidad seguidores", "cantidad de publicaciones", 
                            "Me gusta promedio 👍", "Comentarios promedio 💬", 
                            "Tasa de interacción 📊"]
            
            actual_pivot_values = [col for col in pivot_values if col in df_accounts.columns and pd.api.types.is_numeric_dtype(df_accounts[col])]

            if actual_pivot_values:
                pivot_table_accounts = pd.pivot_table(df_accounts, 
                                                    values=actual_pivot_values, 
                                                    index=["Nombre de usuario"],
                                                    aggfunc=np.mean)
                
                pivot_table_accounts.to_excel(writer, sheet_name='Análisis Cuentas', startrow=1, startcol=0)
                worksheet_analysis.write('A1', 'Análisis de Cuentas de Instagram - Resumen de Métricas')

            if not df_posts.empty:
                worksheet_post_analysis = workbook.add_worksheet('Análisis Posts')
                
                pivot_table_posts = pd.pivot_table(df_posts, 
                                                values=["Likes", "Comments"],
                                                index=["Nombre de usuario", "Type"],
                                                aggfunc=np.mean)
                
                pivot_table_posts.to_excel(writer, sheet_name='Análisis Posts', startrow=1, startcol=0)
                worksheet_post_analysis.write('A1', 'Análisis de Posts de Instagram por Tipo y Usuario')
            
            return True

    except Exception as e:
        print(f"Error al exportar a Excel: {e}")
        return False

# --- INTERFAZ GRÁFICA ---

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Instagram Scraper")
        self.geometry("1000x800")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        # Frame de Entrada
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.input_frame.grid_columnconfigure(0, weight=1)
        self.input_frame.grid_columnconfigure(1, weight=0) 
        self.input_frame.grid_columnconfigure(2, weight=0)


        ctk.CTkLabel(self.input_frame, text="Nombres de Usuario de Instagram (separados por comas):").grid(row=0, column=0, padx=10, pady=10, sticky="w", columnspan=2)
        self.usernames_entry = ctk.CTkEntry(self.input_frame, placeholder_text="ej: instagram,natgeo,cristiano", width=400)
        self.usernames_entry.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self.input_frame, text="Posts a obtener:").grid(row=0, column=2, padx=(10,0), pady=10, sticky="w")
        self.posts_count_combobox = ctk.CTkComboBox(self.input_frame, 
                                                    values=["2", "5", "12", "24", "50"],
                                                    command=self.set_posts_count)
        self.posts_count_combobox.set("12")
        self.posts_count_combobox.grid(row=1, column=2, padx=10, pady=10, sticky="e")

        self.start_button = ctk.CTkButton(self.input_frame, text="Iniciar Análisis", command=self.start_scraping)
        self.start_button.grid(row=1, column=1, padx=10, pady=10, sticky="e")


        # Frame de Resultados (consola y detalles)
        self.results_frame = ctk.CTkFrame(self)
        self.results_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        self.results_frame.grid_rowconfigure(0, weight=0)
        self.results_frame.grid_rowconfigure(1, weight=1)
        self.results_frame.grid_columnconfigure(0, weight=1)
        
        # Log de consola
        self.output_log = ctk.CTkTextbox(self.results_frame, width=800, height=100, state="disabled", wrap="word")
        self.output_log.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="nsew")

        # Frame para los detalles
        self.data_display_frame = ctk.CTkScrollableFrame(self.results_frame, label_text="Resultados del Análisis:")
        self.data_display_frame.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew")
        self.data_display_frame.grid_columnconfigure(0, weight=1)

        self.detailed_results_text = ctk.CTkTextbox(self.data_display_frame, height=250, state="disabled", wrap="word")
        self.detailed_results_text.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Botones de Acción (Descargar)
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, padx=20, pady=20, sticky="nsew")
        self.action_frame.grid_columnconfigure((0, 1), weight=1)

        self.download_json_button = ctk.CTkButton(self.action_frame, text="Descargar Resultados (JSON)", command=self.download_json)
        self.download_json_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.download_json_button.configure(state="disabled")

        self.download_excel_button = ctk.CTkButton(self.action_frame, text="Descargar Resultados (Excel con Tablas)", command=self.download_excel_with_charts)
        self.download_excel_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.download_excel_button.configure(state="disabled")

        self.scraped_data = []
        self.selected_posts_count = int(self.posts_count_combobox.get())

    def set_posts_count(self, choice):
        try:
            self.selected_posts_count = int(choice)
            self.update_log(f"Cantidad de posts a obtener configurada a: {self.selected_posts_count}")
        except ValueError:
            self.update_log("Error: Valor inválido para cantidad de posts. Usando 12 por defecto.")
            self.selected_posts_count = 12
            self.posts_count_combobox.set("12")

    def update_log(self, message):
        self.output_log.configure(state="normal")
        self.output_log.insert("end", message + "\n")
        self.output_log.see("end")
        self.output_log.configure(state="disabled")
        self.update_idletasks()

    def start_scraping(self):
        usernames_input = self.usernames_entry.get().strip()
        if not usernames_input:
            messagebox.showwarning("Entrada Vacía", "Por favor, introduce al menos un nombre de usuario.")
            return

        usernames_list = [u.strip() for u in usernames_input.split(',') if u.strip()]
        if not usernames_list:
            messagebox.showwarning("Entrada Vacía", "Por favor, introduce al menos un nombre de usuario válido.")
            return

        self.output_log.configure(state="normal")
        self.output_log.delete("1.0", "end")
        self.output_log.configure(state="disabled")
        self.detailed_results_text.configure(state="normal")
        self.detailed_results_text.delete("1.0", "end")
        self.detailed_results_text.configure(state="disabled")

        self.scraped_data = []
        self.start_button.configure(state="disabled", text="Analizando...")
        self.download_json_button.configure(state="disabled")
        self.download_excel_button.configure(state="disabled")

        Thread(target=self._run_scraping_thread, args=(usernames_list, self.selected_posts_count,)).start()

    def _run_scraping_thread(self, usernames_list, posts_to_fetch):
        try:
            self.update_log("Iniciando análisis de Instagram...")
            results = scrape_instagram_profiles(usernames_list, posts_to_fetch, self._report_progress)
            self.scraped_data = results
            
            self.after(0, self._display_final_results, results)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error de Scraping", f"Ocurrió un error inesperado: {e}"))
            self.update_log(f"Error fatal: {e}")
        finally:
            self.after(0, lambda: self.start_button.configure(state="normal", text="Iniciar Análisis"))
            self.after(0, lambda: self.download_json_button.configure(state="normal" if self.scraped_data else "disabled"))
            self.after(0, lambda: self.download_excel_button.configure(state="normal" if self.scraped_data else "disabled"))

    def _report_progress(self, type, message):
        if type == "log":
            self.update_log(message)
        elif type == "results":
            pass

    def _display_final_results(self, results):
        if not results:
            self.update_log("No se pudieron obtener datos para ningún usuario.")
            return

        detailed_output = ""
        for user in results:
            if user.get("Error al obtener perfil") == "Usuario no encontrado":
                detailed_output += f"--- {user.get('Nombre de usuario', 'Usuario Desconocido')} ---\n"
                detailed_output += "El usuario ingresado no existe.\n\n"
                continue
                
            detailed_output += f"--- {user.get('Nombre de usuario', 'Usuario Desconocido')} ---\n"
            detailed_output += f"Nombre completo: {user.get('Nombre completo', 'N/A')}\n"
            detailed_output += f"Biografía: {user.get('Biografía', 'N/A')}\n"
            detailed_output += f"País: {user.get('País', 'N/A')}\n"
            detailed_output += f"cantidad seguidores: {user.get('cantidad seguidores', 'N/A')}\n"
            detailed_output += f"cantidad seguidos: {user.get('cantidad seguidos', 'N/A')}\n"
            detailed_output += f"cantidad de publicaciones: {user.get('cantidad de publicaciones', 'N/A')}\n"
            detailed_output += f"Está verificado ✅: {user.get('Está verificado ✅', 'N/A')}\n"
            detailed_output += f"¿Es una cuenta profesional?: {user.get('¿Es una cuenta profesional?', 'N/A')}\n"
            
            # --- CAMBIO AQUÍ: Añadir la línea del promedio de posts ---
            posts_for_avg = user.get('Posts para promedio', 'N/A')
            if posts_for_avg != 'N/A' and posts_for_avg > 0:
                detailed_output += f"----- Promedio de los últimos {posts_for_avg} posts -----\n"
            else:
                detailed_output += f"----- Promedio de posts (No Disponible) -----\n"

            detailed_output += f"Me gusta promedio 👍: {user.get('Me gusta promedio 👍', 'N/A')}\n"
            detailed_output += f"Comentarios promedio 💬: {user.get('Comentarios promedio 💬', 'N/A')}\n"
            detailed_output += f"Tasa de interacción 📊: {user.get('Tasa de interacción 📊', 'N/A')}\n"

            if "Error al obtener perfil" in user and user.get("Error al obtener perfil") != "Usuario no encontrado":
                detailed_output += f"Error al obtener perfil: {user.get('Error al obtener perfil', 'N/A')}\n"
            if "Error al obtener medios" in user:
                detailed_output += f"Error al obtener medios: {user.get('Error al obtener medios', 'N/A')}\n"

            detailed_output += "\n"

        self.detailed_results_text.configure(state="normal")
        self.detailed_results_text.delete("1.0", "end")
        self.detailed_results_text.insert("end", detailed_output)
        self.detailed_results_text.configure(state="disabled")


    def download_json(self):
        if not self.scraped_data:
            messagebox.showwarning("No hay datos", "No hay datos para descargar. Realiza un scraping primero.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile="instagram_user_data_latest.json"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.scraped_data, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("Descarga Exitosa", f"Datos guardados en:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error al guardar JSON", f"No se pudo guardar el archivo JSON:\n{e}")

    def download_excel_with_charts(self):
        if not self.scraped_data:
            messagebox.showwarning("No hay datos", "No hay datos para descargar. Realiza un scraping primero.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="instagram_user_data_analysis.xlsx"
        )
        if file_path:
            try:
                success = export_to_excel_with_pivot_and_charts(self.scraped_data, file_path)
                if success:
                    messagebox.showinfo("Descarga Exitosa", f"Datos con tablas guardados en:\n{file_path}")
                else:
                    messagebox.showerror("Error al guardar Excel", "No se pudo guardar el archivo Excel. Revisa el log para más detalles.")
            except Exception as e:
                messagebox.showerror("Error al guardar Excel", f"Ocurrió un error al intentar guardar el archivo Excel:\n{e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()