package main

import "testing"

func TestGetCord(t *testing.T) {
	cord, err := GetCordinate(2, 3)
	if err != nil {
		t.Fatalf("got err: %s", err.Error())
	}

	if cord != "B3" {
		t.Fatalf("expected B3 got %s", cord)
	}
}

func TestGetCordinateRange(t *testing.T) {
	cord, err := GetCordinateRange("Test", 1, 2, 8)
	if err != nil {
		t.Fatalf("got err: %s", err.Error())
	}

	if cord != "Test!$A$2:$A$8" {
		t.Fatalf("expected 'Test!$A$2:$A$8' got %s", cord)
	}
}
