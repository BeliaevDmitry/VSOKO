package org.school.analysis.repository.impl;

import org.school.analysis.model.StudentResult;
import org.school.analysis.repository.StudentResultRepository;

import java.util.ArrayList;
import java.util.List;

public class InMemoryStudentRepository implements StudentResultRepository {

    private final List<StudentResult> storage = new ArrayList<>();

    @Override
    public boolean save(StudentResult studentResult) {
        return storage.add(studentResult);
    }

    @Override
    public int saveAll(List<StudentResult> studentResults) {
        int count = 0;
        for (StudentResult result : studentResults) {
            if (save(result)) {
                count++;
            }
        }
        return count;
    }

    @Override
    public boolean exists(StudentResult studentResult) {
        return storage.stream()
                .anyMatch(s -> s.getFio().equals(studentResult.getFio()) &&
                        s.getClassName().equals(studentResult.getClassName()) &&
                        s.getSubject().equals(studentResult.getSubject()));
    }

    public List<StudentResult> getAll() {
        return new ArrayList<>(storage);
    }

    public void clear() {
        storage.clear();
    }
}