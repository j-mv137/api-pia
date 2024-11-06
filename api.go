package main

import (
	"encoding/json"
	"fmt"
	"math/rand"
	"net/http"
	"os"
	"strconv"

	"github.com/gorilla/mux"
)

func (s *APIServer) Run() {
	router := mux.NewRouter()
	router.HandleFunc("/api/new-survey", makeHTTPHandlerFunc(s.handleNewSurvey))
	router.HandleFunc("/api/register", makeHTTPHandlerFunc(s.handleRegister))
	router.HandleFunc("/api/calif", makeHTTPHandlerFunc(s.handleNewGrade))
	router.HandleFunc("/api/download-xlsx", makeHTTPHandlerFunc(s.handleDownloadFile))

	fmt.Printf("Server running on port http://localhost%s", s.ListenAddr)

	http.ListenAndServe(s.ListenAddr, corsMiddleware(router))
}

func corsMiddleware(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Add("Access-Control-Allow-Origin", "*")                   // Permite solo este origen
		w.Header().Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS") // Métodos permitidos
		w.Header().Add("Access-Control-Allow-Headers", "Content-Type")       // Encabezados permitidos

		// Permite que las solicitudes OPTIONS se resuelvan sin pasar al siguiente controlador
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusNoContent)
			return
		}

		next.ServeHTTP(w, r)
	})
}
func (s *APIServer) handleNewGrade(w http.ResponseWriter, r *http.Request) error {
	var id int
	sheet := "Registro de calificaciones"

	grade := &GradeForm{}
	err := json.NewDecoder(r.Body).Decode(grade)

	if err != nil {
		return err
	}

	gradeInt, err := strconv.ParseInt(grade.Calif, 10, 32)

	if err != nil {
		fmt.Println(err)
		return err
	}

	rows, err := s.ExcelFile.GetRows("Registro de calificaciones")
	fmt.Println(len(rows))
	if err != nil {
		return err
	}

	for i := 0; i < len(rows); i++ {
		if i < 2 {
			continue
		}
		cellVal, err := s.ExcelFile.GetCellValue(sheet, "A"+strconv.FormatInt(int64(i), 10))
		if err != nil {
			break
		}
		if len(cellVal) == 0 {
			id = i
			break
		}
	}

	idStr := strconv.FormatInt(int64(id), 10)
	err1 := s.ExcelFile.SetCellValue(sheet, "A"+idStr, id-2)
	if err1 != nil {
		return err1
	}
	err2 := s.ExcelFile.SetCellValue(sheet, "B"+idStr, gradeInt)
	if err2 != nil {
		return err2
	}

	if err := s.ExcelFile.Save(); err != nil {
		return err
	}
	defer s.ExcelFile.Close()
	return nil
}

func (s *APIServer) handleNewSurvey(w http.ResponseWriter, r *http.Request) error {
	formData := &FormType{}
	err := json.NewDecoder(r.Body).Decode(formData)
	var id int

	if err != nil {
		return err
	}

	userResponse, err := makeUserResponse(*formData)

	if err != nil {
		return err
	}

	rows, err := s.ExcelFile.GetRows("IPAQ Short Form Scoring")

	if err != nil {
		return err
	}

	for i := 0; i < len(rows); i++ {
		if i < 7 {
			continue
		}
		cellVal, err := s.ExcelFile.GetCellValue("IPAQ Short Form Scoring", "A"+strconv.FormatInt(int64(i), 10))
		if err != nil {
			break
		}
		if len(cellVal) == 0 {
			id = i
			break
		}
	}

	userResponse.id = id - 7
	idStr := strconv.FormatInt(int64(id), 10)

	err2 := s.WriteExcelIPAQ("IPAQ Short Form Scoring", idStr, userResponse)
	if err2 != nil {
		return err2
	}

	return nil
}

func (s *APIServer) handleRegister(w http.ResponseWriter, r *http.Request) error {
	userRegister := &UserRegister{}
	err := json.NewDecoder(r.Body).Decode(userRegister)
	var id int

	if err != nil {
		return err
	}

	rows, err := s.ExcelFile.GetRows("Registro de usuarios")
	fmt.Println(len(rows))
	if err != nil {
		return err
	}

	for i := 0; i < len(rows); i++ {
		if i < 2 {
			continue
		}
		cellVal, err := s.ExcelFile.GetCellValue("Registro de usuarios", "A"+strconv.FormatInt(int64(i), 10))
		if err != nil {
			break
		}
		if len(cellVal) == 0 {
			id = i
			break
		}
	}
	fmt.Println(id)
	userRegisterC, err := makeUserRegister(*userRegister, id)

	if err != nil {
		return err
	}

	errExcelReg := s.WriteExcelRegister("Registro de usuarios", id, userRegisterC)

	if errExcelReg != nil {
		return errExcelReg
	}

	return nil
}

func (s *APIServer) WriteExcelRegister(sheet string, id int, userRegisterC *UserRegisterC) error {
	idStr := strconv.FormatInt(int64(id), 10)
	err1 := s.ExcelFile.SetCellValue(sheet, "A"+idStr, userRegisterC.ID)
	if err1 != nil {
		return err1
	}
	err2 := s.ExcelFile.SetCellValue(sheet, "B"+idStr, userRegisterC.Sexo)
	if err2 != nil {
		return err2
	}
	err3 := s.ExcelFile.SetCellValue(sheet, "C"+idStr, userRegisterC.Edad)
	if err3 != nil {
		return err3
	}
	err4 := s.ExcelFile.SetCellValue(sheet, "D"+idStr, userRegisterC.Bachillerato)
	if err4 != nil {
		return err4
	}
	err5 := s.ExcelFile.SetCellValue(sheet, "E"+idStr, userRegisterC.Semestre)
	if err5 != nil {
		return err5
	}
	err6 := s.ExcelFile.SetCellValue(sheet, "F"+idStr, userRegisterC.EstadoCivil)
	if err6 != nil {
		return err6
	}
	err7 := s.ExcelFile.SetCellValue(sheet, "G"+idStr, userRegisterC.Trabajo)
	if err7 != nil {
		return err7
	}
	err8 := s.ExcelFile.SetCellValue(sheet, "H"+idStr, userRegisterC.Etnia)
	if err8 != nil {
		return err8
	}

	if err := s.ExcelFile.Save(); err != nil {
		return err
	}
	defer s.ExcelFile.Close()
	return nil
}

func (s *APIServer) WriteExcelIPAQ(sheet string, idStr string, userResponse *UserResponse) error {
	err1 := s.ExcelFile.SetCellValue(sheet, "A"+idStr, userResponse.id)
	if err1 != nil {
		return err1
	}
	err2 := s.ExcelFile.SetCellValue(sheet, "B"+idStr, userResponse.Q1)
	if err2 != nil {
		return err2
	}
	err3 := s.ExcelFile.SetCellValue(sheet, "C"+idStr, userResponse.Q2)
	if err3 != nil {
		return err3
	}
	err4 := s.ExcelFile.SetCellValue(sheet, "D"+idStr, userResponse.Q3)
	if err4 != nil {
		return err4
	}
	err5 := s.ExcelFile.SetCellValue(sheet, "E"+idStr, userResponse.Q4)
	if err5 != nil {
		return err5
	}
	err6 := s.ExcelFile.SetCellValue(sheet, "F"+idStr, userResponse.Q5)
	if err6 != nil {
		return err6
	}
	err7 := s.ExcelFile.SetCellValue(sheet, "G"+idStr, userResponse.Q6)
	if err7 != nil {
		return err7
	}
	err8 := s.ExcelFile.SetCellValue(sheet, "H"+idStr, userResponse.Q7)
	if err8 != nil {
		return err8
	}

	if err := s.ExcelFile.Save(); err != nil {
		return err
	}
	defer s.ExcelFile.Close()
	return nil
}

func (s *APIServer) handleDownloadFile(w http.ResponseWriter, r *http.Request) error {
	filepath := os.Getenv("EXCEL_FILE_PATH")

	file, err := os.Open(filepath)

	if err != nil {
		return err
	}

	fileInfo, err := file.Stat()

	if err != nil {
		return err
	}

	w.Header().Add("Content-Disposition", "attachment; filename="+filepath)
	w.Header().Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	http.ServeContent(w, r, filepath, fileInfo.ModTime(), file)

	return nil
}

func makeUserRegister(user UserRegister, id int) (*UserRegisterC, error) {
	var sem, age int64
	var w, e bool

	switch user.Semestre {
	case "Primero":
		sem = 1
	case "Segundo":
		sem = 2
	case "Tercero":
		sem = 3
	case "Cuarto":
		sem = 4
	case "Quinto":
		sem = 5

	case "Sexto":
		sem = 6
	}

	age, err := strconv.ParseInt(user.Edad, 10, 32)

	if err != nil {
		return &UserRegisterC{}, err
	}

	switch user.Trabajo {
	case "Si":
		w = true
	case "No":
		w = false
	}

	switch user.Etnia {
	case "Si":
		e = true
	case "No":
		e = false
	}

	return &UserRegisterC{
		ID:           id - 2,
		Bachillerato: user.Bachillerato,
		Semestre:     int(sem),
		Sexo:         user.Sexo,
		Trabajo:      w,
		Etnia:        e,
		Edad:         int(age),
		EstadoCivil:  user.EstadoCivil,
	}, nil
}

func makeUserResponse(data FormType) (*UserResponse, error) {
	var q1, q2, q3, q4, q5, q6, q7 float64

	if data.Q1R {
		q1 = 0.0
		q2 = 0.0
	} else {
		q1Float, err := strconv.ParseFloat(data.Q1T1, 64)
		if err != nil {
			fmt.Printf("1")
			return &UserResponse{}, err
		}
		q1 = q1Float

		if data.Q2R {
			q2 = 0.0
		} else {
			q2Hr, err := strconv.ParseFloat(data.Q2T1, 64)
			if err != nil {
				fmt.Printf("2h")
				return &UserResponse{}, err
			}
			q2min, err := strconv.ParseFloat(data.Q2T2, 64)
			if err != nil {
				fmt.Printf("2m")
				return &UserResponse{}, err
			}

			q2 = q2Hr*60 + q2min
		}
	}

	if data.Q3R {
		q3 = 0.0
		q4 = 0.0
	} else {
		q3Float, err := strconv.ParseFloat(data.Q3T1, 64)
		if err != nil {
			fmt.Printf("3")
			return &UserResponse{}, err
		}
		q3 = q3Float

		if data.Q4R {
			q4 = 0.0
		} else {
			q4Hr, err := strconv.ParseFloat(data.Q4T1, 64)
			if err != nil {
				fmt.Printf("4h")
				return &UserResponse{}, err
			}
			q4min, err := strconv.ParseFloat(data.Q4T2, 64)
			if err != nil {
				fmt.Printf("%s", data.Q4T2)
				return &UserResponse{}, err
			}

			q4 = q4Hr*60 + q4min
		}
	}

	if data.Q5R {
		q5 = 0.0
		q6 = 0.0
	} else {
		q5Float, err := strconv.ParseFloat(data.Q5T1, 64)
		q5 = q5Float
		if err != nil {
			fmt.Printf("5")
			return &UserResponse{}, err
		}

		if data.Q6R {
			q6 = 0.0
		} else {
			q6Hr, err := strconv.ParseFloat(data.Q6T1, 64)
			if err != nil {
				fmt.Printf("6h")
				return &UserResponse{}, err
			}
			q6min, err := strconv.ParseFloat(data.Q6T2, 64)
			if err != nil {
				fmt.Printf("6m")
				return &UserResponse{}, err
			}

			q6 = q6Hr*60 + q6min
		}
	}

	if data.Q7R {
		q7 = 0.0
	} else {
		q7Float, err := strconv.ParseFloat(data.Q7T1, 64)

		if err != nil {
			fmt.Printf("7")
			return &UserResponse{}, err
		}

		q7 = q7Float
	}

	return &UserResponse{
		id: rand.Intn(10000),
		Q1: q1,
		Q2: q2,
		Q3: q3,
		Q4: q4,
		Q5: q5,
		Q6: q6,
		Q7: q7,
	}, nil
}

func makeHTTPHandlerFunc(f APIHandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		if err := f(w, r); err != nil {
			WriteJSON(w, http.StatusBadRequest, APIError{Error: err.Error()})
		}
	}
}

func WriteJSON(w http.ResponseWriter, status int, v any) error {
	w.Header().Add("Content-Type", "application/json")
	// w.Header().Add("Access-Control-Allow-Origin", "*")
	w.WriteHeader(status)
	return json.NewEncoder(w).Encode(v)
}
