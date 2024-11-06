package main

import (
	"net/http"

	"github.com/xuri/excelize/v2"
)

type APIHandlerFunc func(http.ResponseWriter, *http.Request) error

type APIServer struct {
	ListenAddr string
	ExcelFile  *excelize.File
}

type FormType struct {
	Q1T1 string `json:"q1T1"`
	Q1R  bool   `json:"q1R"`
	Q2T1 string `json:"q2T1"`
	Q2T2 string `json:"q2T2"`
	Q2R  bool   `json:"q2R"`
	Q3T1 string `json:"q3T1"`
	Q3R  bool   `json:"q3R"`
	Q4T1 string `json:"q4T1"`
	Q4T2 string `json:"q4T2"`
	Q4R  bool   `json:"q4R"`
	Q5T1 string `json:"q5T1"`
	Q5R  bool   `json:"q5R"`
	Q6T1 string `json:"q6T1"`
	Q6T2 string `json:"q6T2"`
	Q6R  bool   `json:"q6R"`
	Q7T1 string `json:"q7T1"`
	Q7R  bool   `json:"q7R"`
}

type UserRegister struct {
	Bachillerato string `json:"bachillerato"`
	Semestre     string `json:"semestre"`
	Sexo         string `json:"sexo"`
	Edad         string `json:"edad"`
	EstadoCivil  string `json:"estadoCivil"`
	Trabajo      string `json:"trabajo"`
	Etnia        string `json:"etnia"`
}
type GradeForm struct {
	Calif string `json:"calif"`
}
type UserRegisterC struct {
	ID           int
	Bachillerato string
	Semestre     int
	Sexo         string
	Edad         int
	EstadoCivil  string
	Trabajo      bool
	Etnia        bool
}

type UserResponse struct {
	id int
	Q1 float64
	Q2 float64
	Q3 float64
	Q4 float64
	Q5 float64
	Q6 float64
	Q7 float64
}

func NewAPIServer(listenAddr string, f *excelize.File) *APIServer {
	return &APIServer{
		ListenAddr: listenAddr,
		ExcelFile:  f,
	}
}

type APIError struct {
	Error string `json:"error"`
}
