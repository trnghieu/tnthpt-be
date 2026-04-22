import mongoose from 'mongoose';
import { SUBJECT_KEYS } from '../constants/subjects.js';

const scoresSchemaDefinition = SUBJECT_KEYS.reduce((accumulator, key) => {
  accumulator[key] = {
    type: Number,
    min: 0,
    max: 10,
    default: null
  };
  return accumulator;
}, {});

const candidateSchema = new mongoose.Schema(
  {
    examNumber: {
      type: String,
      required: true,
      unique: true,
      trim: true,
      index: true
    },
    fullName: {
      type: String,
      required: true,
      trim: true
    },
    schoolName: {
      type: String,
      default: '',
      trim: true
    },
    examSubjects: {
      type: String,
      default: 'Toán, Lý, Hóa',
      trim: true
    },
    examRoom: {
      type: String,
      default: '',
      trim: true
    },
    scores: scoresSchemaDefinition
  },
  {
    timestamps: true,
    toJSON: {
      transform: (_, returnedObject) => {
        delete returnedObject.__v;
        return returnedObject;
      }
    }
  }
);

const Candidate = mongoose.model('Candidate', candidateSchema);
export default Candidate;
