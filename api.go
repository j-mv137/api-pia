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

	fmt.Println(rows[7][1])

	for i, _ := range rows {
		if i < 7 {
			continue
		}
		cellVal, err := s.ExcelFile.GetCellValue("IPAQ Short Form Scoring", "A"+strconv.FormatInt(int64(i), 10))
		if err != nil {
			break
		}
		if len(cellVal) == 0 {
			id = i
			fmt.Println(id)
			break
		}
	}

	userResponse.id = id
	idStr := strconv.FormatInt(int64(userResponse.id), 10)

	err2 := s.WriteExcel("IPAQ Short Form Scoring", idStr, userResponse)
	if err2 != nil {
		return err2
	}

	return nil
}

func (s *APIServer) WriteExcel(sheet string, idStr string, userResponse *UserResponse) error {
	// idInt, err := strconv.ParseInt(idStr, 10, 64)

	// if err != nil {
	// 	return err
	// }

	err1 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "A"+idStr, idStr)
	if err1 != nil {
		return err1
	}
	err2 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "B"+idStr, userResponse.Q1)
	if err2 != nil {
		return err2
	}
	err3 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "C"+idStr, userResponse.Q2)
	if err3 != nil {
		return err3
	}
	err4 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "D"+idStr, userResponse.Q3)
	if err4 != nil {
		return err4
	}
	err5 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "E"+idStr, userResponse.Q4)
	if err5 != nil {
		return err5
	}
	err6 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "F"+idStr, userResponse.Q5)
	if err6 != nil {
		return err6
	}
	err7 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "G"+idStr, userResponse.Q6)
	if err7 != nil {
		return err7
	}
	err8 := s.ExcelFile.SetCellValue("IPAQ Short Form Scoring", "H"+idStr, userResponse.Q7)
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
